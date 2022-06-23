<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 
Session.LCID = 1046
nivel=4

pagina=request.QueryString("pagina")

if pagina>=1 then
	ori = session("ori_age")
	opt = session("opt_age")
	session("ori_age")= ori
	session("opt_age")=	opt	
else
	ori = request.QueryString("or")
	opt= request.QueryString("opt")
	session("ori_age")= ori
	session("opt_age")=	opt
end if	

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
trava=session("trava")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

if trava="n" AND (ori=1 or ori="1")then

	cod= request.QueryString("cod_cons")
	ordem= session("ordem")
	status_entrevista=session("status_entrevista")
	
	dia_de= session("dia_de")
	mes_de= session("mes_de")
	ano_de= session("ano_de")
	hora_de= session("hora_de")
	min_de= session("min_de")
	hora_imp_de=hora_de
	
	data_de=mes_de&"/"&dia_de&"/"&ano_de
	
	
	dia_de=dia_de*1
	mes_de=mes_de*1
'	h_de=hora_de*1
'	min_de=min_de*1
	
	if dia_de<10 then
	dia_de="0"&dia_de
	end if
	if mes_de<10 then
	mes_de="0"&mes_de
	end if
'	if h_de<10 then
'	h_de="0"&h_de
'	end if
'	if min_de<10 then
'	min_de="0"&min_de
'	end if
'	
'	hora_de=h_de&":"&min_de
	
	'data_inicio=dia_de&"/"&mes_de&"/"&ano_de&", "&hora_de
	data_inicio=dia_de&"/"&mes_de&"/"&ano_de	
	
	dia_ate= session("dia_ate")
	mes_ate= session("mes_ate")
	ano_ate= session("ano_ate")
	hora_ate= session("hora_ate")
	min_ate= session("min_ate")
	
	hora_imp_ate=hora_ate
	
	data_ate=mes_ate&"/"&dia_ate&"/"&ano_ate
	
	dia_ate=dia_ate*1
	mes_ate=mes_ate*1
'	h_ate=hora_ate*1
'	min_ate=min_ate*1
	
	if dia_ate<10 then
	dia_ate="0"&dia_ate
	end if
	if mes_ate<10 then
	mes_ate="0"&mes_ate
	end if
'	if h_ate<10 then
'	h_ate="0"&h_ate
'	end if
'	if min_ate<10 then
'	min_ate="0"&min_ate
'	end if
'	
'	hora_ate=h_ate&":"&min_ate	
	'data_fim=dia_ate&"/"&mes_ate&"/"&ano_ate&", "&hora_ate
	data_fim=dia_ate&"/"&mes_ate&"/"&ano_ate


elseif trava="n" AND (ori=2 or ori="2")then
	'ordem= "dt"
	obr=session("obr")
	session("obr")=obr
	'obr=cod&"?"&ordem&"?"&status_entrevista&"?"&data_de&"?"&hora_de&"?"&data_inicio&"?"&data_ate&"?"&hora_ate&"?"&data_fim
	'response.Write(">>"& obr)
	dados= split(obr, "?" )
	cod= dados(0)
	ordem= dados(1)
	status_entrevista= dados(2)
	data_de= dados(3)
	hora_de= dados(4)
	data_inicio= dados(5)
	data_ate= dados(6)
	hora_ate= dados(7)
	data_fim= dados(8)
	
	dados_dtd= split(data_de, "/" )
	dia_de= dados_dtd(0)
	mes_de= dados_dtd(1)
	ano_de= dados_dtd(2)
	
	
	
	dados_hrd= split(hora_de, ":" )
'	h_de= dados_hrd(0)
'	min_de= dados_hrd(1)
	
	dados_dta= split(data_ate, "/" )
	dia_ate= dados_dta(0)
	mes_ate= dados_dta(1)
	ano_ate= dados_dta(2)
	
	dados_hra= split(hora_ate, ":" )
'	h_ate= dados_hra(0)
'	min_ate= dados_hra(1)
elseif trava="n" AND (ori=3 or ori="3")then
	cod= request.form("cod")
	ordem= request.form("ordem")
	status_entrevista=request.form("status")
	data_de= request.form("data_de")
	hora_de= request.form("hora_de")
	data_inicio= request.form("data_inicio")
	data_ate= request.form("data_ate")
	hora_ate= request.form("hora_ate")
	data_fim= request.form("data_fim")
	
	
	
	dados_dtd= split(data_de, "/" )
	dia_de= dados_dtd(0)
	mes_de= dados_dtd(1)
	ano_de= dados_dtd(2)
	
	dados_hrd= split(hora_de, ":" )
'	h_de= dados_hrd(0)
'	min_de= dados_hrd(1)
	
	dados_dta= split(data_ate, "/" )
	dia_ate= dados_dta(0)
	mes_ate= dados_dta(1)
	ano_ate= dados_dta(2)
	
	
	dados_hra= split(hora_ate, ":" )
'	h_ate= dados_hra(0)
'	min_ate= dados_hra(1)





else
	cod= request.form("cod")
	ordem= request.form("ordem")
	status_entrevista=request.form("status")
	
	dia_de= request.form("dia_de")
	mes_de= request.form("mes_de")
	ano_de= request.form("ano_de")
	hora_de= request.form("hora_de")
	min_de= request.form("min_de")
	hora_imp_de=hora_de
	
	data_de=mes_de&"/"&dia_de&"/"&ano_de
	
	
	dia_de=dia_de*1
	mes_de=mes_de*1
	h_de=hora_de*1
	min_de=min_de*1
	
	if dia_de<10 then
	dia_de="0"&dia_de
	end if
	if mes_de<10 then
	mes_de="0"&mes_de
	end if
	if h_de<10 then
	h_de="0"&h_de
	end if
	if min_de<10 then
	min_de="0"&min_de
	end if
	
	hora_de=h_de&":"&min_de
	
	'data_inicio=dia_de&"/"&mes_de&"/"&ano_de&", "&hora_de
	data_inicio=dia_de&"/"&mes_de&"/"&ano_de	
	
	dia_ate= request.form("dia_ate")
	mes_ate= request.form("mes_ate")
	ano_ate= request.form("ano_ate")
	hora_ate= request.form("hora_ate")
	min_ate= request.form("min_ate")
	
	hora_imp_ate=hora_ate
	
	data_ate=mes_ate&"/"&dia_ate&"/"&ano_ate
	
	dia_ate=dia_ate*1
	mes_ate=mes_ate*1
	h_ate=hora_ate*1
	min_ate=min_ate*1
	
	if dia_ate<10 then
	dia_ate="0"&dia_ate
	end if
	if mes_ate<10 then
	mes_ate="0"&mes_ate
	end if
	if h_ate<10 then
	h_ate="0"&h_ate
	end if
	if min_ate<10 then
	min_ate="0"&min_ate
	end if
	
	hora_ate=h_ate&":"&min_ate	
	'data_fim=dia_ate&"/"&mes_ate&"/"&ano_ate&", "&hora_ate
	data_fim=dia_ate&"/"&mes_ate&"/"&ano_ate
end if

sessionobr=cod&"?"&ordem&"?"&status_entrevista&"?"&data_de&"?"&hora_de&"?"&data_inicio&"?"&data_ate&"?"&hora_ate&"?"&data_fim
trava=session("trava")
ocorr= request.form("ocorr")
session("obr")=sessionobr
'Para o arquivo de impressão
obr=cod&"?"&ordem&"?"&status_entrevista&"?"&dia_de&"?"&mes_de&"?"&ano_de&"?"&h_de&"?"&min_de&"?"&dia_ate&"?"&mes_ate&"?"&ano_ate&"?"&h_ate&"?"&min_ate

dados_msg  = cod&"?"&ordem&"?"&status_entrevista&"?"&data_de&"?"&hora_de&"?"&data_inicio&"?"&data_ate&"?"&hora_ate&"?"&data_fim

if status_entrevista="" or isnull(status_entrevista) then
	status_entrevista_form = "Todos"
else
	entrevistas = split(status_entrevista,",")
	for s = 0 to ubound(entrevistas)
		Select case entrevistas(s)		
			case 1
			nome_status="Atendida"
			
			case 2
			nome_status="Cancelada"
			
			case 3
			nome_status="Pendente"	
		end select	
		if s = 0 then
			status_entrevista_form = nome_status	
		else
			status_entrevista_form = status_entrevista_form&", "&nome_status
		end if		
	next		
end if


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHO_e & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4		
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CONp = Server.CreateObject("ADODB.Connection") 
		ABRIRp = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONp.Open ABRIRp		
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	

if cod=0 then
	no_aluno = "Todos"
	no_unidade= "Todas"
	no_curso="Todos"	
	no_etapa="Todas" 
	turma ="Todas"		
	cod_link = 0
	permite_incluir = "N"	
	
else		
	permite_incluir = "S"	
	
'	Set RS = Server.CreateObject("ADODB.Recordset")
'	SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
'	RS.Open SQL, CON1
		
		





	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Alunos,TB_Matriculas WHERE TB_Alunos.CO_Matricula = TB_Matriculas.CO_Matricula AND NU_Ano="& ano_letivo &" AND TB_Matriculas.CO_Matricula ="& cod
	RS.Open SQL, CON1

	no_aluno = RS("NO_Aluno")
	ano_aluno = RS("NU_Ano")
	rematricula = RS("DA_Rematricula")
	situacao = RS("CO_Situacao")
	encerramento= RS("DA_Encerramento")
	unidade= RS("NU_Unidade")
	curso= RS("CO_Curso")
	etapa= RS("CO_Etapa")
	turma= RS("CO_Turma")
	cham= RS("NU_Chamada")
	
	no_unidade= GeraNomes("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
	no_curso=GeraNomes("C",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
	no_etapa=GeraNomes("E",curso,etapa,variavel3,variavel4,variavel5,CON0,outro) 	
	
	Set RSCONTST = Server.CreateObject("ADODB.Recordset")
	SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
	RSCONTST.Open SQLCONTST, CON0
						
	no_situacao = RSCONTST("TX_Descricao_Situacao")			

	cod_link = cod
end if


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

function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
function checkTheBox() {
   var chk = document.getElementsByName('entrevista')
    var len = chk.length

    for(i=0;i<len;i++)
    {
         if(chk[i].checked){
        return true;
          }
    }
	alert("Pelo menos uma entrevista deve ser selecionada!")		
    return false;
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
<%	 if opt="ok" then    %>    
            <tr>    
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,320,2,0) %>
    </td>
			  </tr>			  			  
<%
elseif opt="ok1" then    %>    
            <tr>    
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,318,2,0) %>
    </td>
			  </tr>
<% end if	%>
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,645,0,"R17") %>
    </td>
			  </tr>			  
          <tr>
      
    <td height="544" valign="top"> 
      <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo"
>
        <tr> 
          <td width="653" class="tb_tit"
>Dados Escolares</td>
          <td width="113" class="tb_tit"
> </td>
        </tr>
        <tr> 
          <td height="10"> <table width="100%" border="0" cellspacing="0">
              <tr> 
                <td width="19%" height="10"> <div align="right"><font class="form_dado_texto"> 
                    Matr&iacute;cula: </font></div></td>
                <td width="9%" height="10"><font class="form_dado_texto"> 
                  <%
				  if cod>0 then
				  	response.Write(cod)
				  end if%>
                  </font></td>
                <td width="6%" height="10"> <div align="right"><font class="form_dado_texto"> 
                    Nome: </font></div></td>
                <td width="66%" height="10"><font class="form_dado_texto"> 
                  <%response.Write(no_aluno)%>
                  </font></td>
              </tr>
            </table></td>
          <td valign="top">&nbsp; </td>
        </tr>
        <tr> 
          <td height="10" bgcolor="#FFFFFF">&nbsp;</td>
          <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
        </tr><form action="resumo.asp?or=3" method="post" name="inativo" id="inativo">
        <tr> 
          <td colspan="2"> 
              <table width="100%" border="0" cellpadding="0" cellspacing="0">
                <tr class="tb_subtit"> 
                  <td width="100" height="10"> <div align="center"> 
                  Ano</div></td>
                  <td width="100" height="10"> <div align="center">Matr&iacute;cula</div></td>
                  <td width="100" height="10"> <div align="center">Cancelamento</div></td>
                  <td width="100" height="10"> <div align="center"> Situa&ccedil;&atilde;o</div></td>
                  <td width="150" height="10"> <div align="center">Unidade</div></td>
                  <td width="150" height="10"> <div align="center">Curso</div></td>
                  <td width="150" height="10"> <div align="center"> Etapa</div></td>
                  <td width="150" height="10"> <div align="center">Turma </div> <div align="center"></div></td>
                </tr>
                <tr class="tb_corpo"> 
                  <td width="100" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(ano_aluno)%>
                      </font></div></td>
                  <td width="100" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(rematricula)%>
                      </font></div></td>
                  <td width="100" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(encerramento)%>
                      </font></div></td>
                  <td width="100" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%					
					response.Write(no_situacao)%>
                      </font></div></td>
                  <td width="150" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_unidade)%>
                      </font></div></td>
                  <td width="150" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_curso)%>
                      </font></div></td>
                  <td width="150" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_etapa)%>
                      </font></div></td>
                  <td width="150" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      </font></div> <div align="center"></div></td>
                </tr>
                <tr> 
                  <td height="10" colspan="8">&nbsp;</td>
                </tr>
                <tr class="tb_tit"> 
                  <td height="10" colspan="8"> Crit&eacute;rios Informados
                      <input name="cod" type="hidden" id="cod" value="<%=cod%>">
                  <input name="nome" type="hidden" class="textInput" id="nome"  value="<%response.Write(no_aluno)%>" size="75" maxlength="50">
                    <input name="status_entrevista" type="hidden" class="textInput" id="status_entrevista"  value="<%response.Write(status_entrevista)%>" size="75" maxlength="50">
                    <input name="data_de" type="hidden" class="textInput" id="data_de"  value="<%response.Write(data_de)%>" size="75" maxlength="50">
                    <input name="hora_de" type="hidden" class="textInput" id="hora_de"  value="<%response.Write(hora_de)%>" size="75" maxlength="50">
                    <input name="data_inicio" type="hidden" class="textInput" id="data_inicio"  value="<%response.Write(data_inicio)%>" size="75" maxlength="50">
                    <input name="data_ate" type="hidden" class="textInput" id="data_ate"  value="<%response.Write(data_ate)%>" size="75" maxlength="50">
                    <input name="hora_ate" type="hidden" class="textInput" id="hora_ate"  value="<%response.Write(hora_ate)%>" size="75" maxlength="50">
                    <input name="data_fim" type="hidden" class="textInput" id="data_fim"  value="<%response.Write(data_fim)%>" size="75" maxlength="50">
                    </td>
                </tr>
                <tr class="tb_subtit"> 
                  <td height="10" colspan="3"><div align="center">Data<!-- e Hora--> 
                  de In&iacute;cio                    </div></td>
                  <td height="10" colspan="2"><div align="center">Data<!-- e Hora -->
                  de Fim</div></td>
                  <td height="10" colspan="2"><div align="center">Status da Entrevista</div></td>
                  <td height="10"> <div align="center">Ordenado por:</div></td>
                </tr>
                <tr class="tb_corpo"> 
                  <td height="10" colspan="3"><div align="center" class="form_dado_texto"><%response.Write(data_inicio)%></div>
                  <div align="center"></div></td>
                  <td height="10" colspan="2"> <div align="center"><font class="form_dado_texto"> 
                      
                      </font><font class="form_dado_texto">
                      <%response.Write(data_fim)%>
                  </font></div></td>
                  <td height="10" colspan="2"><div align="center" class="form_dado_texto"><%response.Write(status_entrevista_form)%></div></td>
                  <td height="10"> <div align="center"><font class="form_dado_texto"> 
                      <select name="ordem" class="select_style"  onChange="MM_callJS('submitfuncao()')">
                        <% if ordem="dt" then%>
                        <option value="dt" selected>Data/Hora</option>
                        <%else%>
                        <option value="dt" >Data/Hora</option>
                        <%end if%>
                        <% if ordem="mt" then%>
                        <option value="mt" selected>Matr&iacute;cula</option>
                        <%else%>
                        <option value="mt" >Matr&iacute;cula</option>
                        <%end if%>
                        <% if ordem="al" then%>
                        <option value="al" selected>Nome Aluno</option>
                        <%else%>
                        <option value="al" >Nome Aluno</option>
                        <%end if%>                                                
                        <% if ordem="en" then%>
                        <option value="en" selected>Tipo de Entrevista</option>
                        <%else%>
                        <option value="en" >Tipo de Entrevista</option>
                        <%end if%>
                        <% if ordem="pr" then%>
                        <option value="pr" selected>Participantes</option>
                        <%else%>
                        <option value="pr" >Participantes</option>
                        <%end if%>
                        <% if ordem="at" then%>
                        <option value="at" selected>Atendido por</option>
                        <%else%>
                        <option value="at" >Atendido por</option>
                        <%end if%>
                        <% if ordem="st" then%>
                        <option value="st" selected>Status</option>
                        <%else%>
                        <option value="st" >Status</option>
                        <%end if%>                        
                      </select>
                  </font></div></td>
                </tr>
              </table>
            </td>
        </tr></form>
        <tr height="10"> 
          <td height="10" colspan="2" >&nbsp;</td>
        </tr>
        <tr> 
          <td height="10" colspan="2" ></td>
        </tr>		
        <tr> 
          <td height="10" colspan="2" class="tb_tit"
>Entrevistas</td>
        </tr>
        <tr > 
          <td height="154" colspan="2" valign="top"> 
           <!-- <form action="redireciona.asp" method="post" name="busca" id="busca" onSubmit="return checksubmit()">--> 		  
 <form action="redireciona.asp" method="post" name="busca" id="busca" onSubmit="return checkTheBox()">		  
		  
              <table width="1000" border="0" cellspacing="0" cellpadding="0">
                <tr class="tb_subtit"> 
                  <td width="21" height="10"> <input type="checkbox" name="todos" class="borda" value="" onClick="this.value=check(this.form.entrevista)"> 
                  </td>
                  <td width="90" align="center">Data / Hora</td>
                  <td width="70" align="center">Matr&iacute;cula</td>
                  <td width="288"> <div align="left">Nome do Aluno<font class="form_dado_texto"> 
                      <input name="cod" type="hidden" id="cod" value="<%=cod%>">
                      <input name="nome" type="hidden" class="textInput" id="nome"  value="<%response.Write(no_aluno)%>" size="75" maxlength="50">
                      <input name="status_entrevista" type="hidden" class="textInput" id="status_entrevista"  value="<%response.Write(status_entrevista)%>" size="75" maxlength="50">
                      <input name="data_de" type="hidden" class="textInput" id="data_de"  value="<%response.Write(data_de)%>" size="75" maxlength="50">
                      <input name="hora_de" type="hidden" class="textInput" id="hora_de"  value="<%response.Write(hora_de)%>" size="75" maxlength="50">
                      <input name="data_inicio" type="hidden" class="textInput" id="data_inicio"  value="<%response.Write(data_inicio)%>" size="75" maxlength="50">
                      <input name="data_ate" type="hidden" class="textInput" id="data_ate"  value="<%response.Write(data_ate)%>" size="75" maxlength="50">
                      </font><font class="form_dado_texto">
                      <input name="hora_ate" type="hidden" class="textInput" id="hora_ate"  value="<%response.Write(hora_ate)%>" size="75" maxlength="50">
                      <input name="data_fim" type="hidden" class="textInput" id="data_fim"  value="<%response.Write(data_fim)%>" size="75" maxlength="50">
                      </font></div></td>
                  <td width="108" align="center">Tipo</td>
                  <td width="122"> <div align="center">Participantes</div></td>
                  <td width="160"> <div align="center">Atendido por</div></td>
                  <td width="159"><div align="center">Status</div></td>
                </tr>
				                <tr> 
                  <td colspan="8"><hr width="1000"></td>
                </tr>
                <%			
cod=cod*1
if cod=0 then
	sql_matricula = ""
else
	sql_matricula = "CO_Matricula ="&cod&" AND "
end if


if status_entrevista="" or isnull(status_entrevista) then
	sql_status_entrevista = ""
else
	sql_status_entrevista = "ST_Entrevista IN("&status_entrevista&") AND "	
end if


Set Rs_ordena = Server.CreateObject ( "ADODB.RecordSet" )
'Vamos adicionar 2 campos nesse recordset!
'O método Append recebe 3 parâmetros:
'Nome do campo, Tipo, Tamanho (opcional)
'O tipo pertence à um DataTypeEnum, e você pode conferir os tipos em
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/ado270/htm/mdcstdatatypeenum.asp
'200 -> VarChar (String), 7 -> Data, 139 -> Numeric
Rs_ordena.Fields.Append "ordem_original", 139, 10
Rs_ordena.Fields.Append "data", 200, 255
Rs_ordena.Fields.Append "data_entrevista", 200, 255
Rs_ordena.Fields.Append "hora_entrevista", 200, 255
Rs_ordena.Fields.Append "co_matric", 139, 10
Rs_ordena.Fields.Append "nome", 200, 255
Rs_ordena.Fields.Append "tipo", 200, 255
Rs_ordena.Fields.Append "participantes", 201, 8000
Rs_ordena.Fields.Append "atendido_por", 200, 255
Rs_ordena.Fields.Append "nome_status", 200, 255

Rs_ordena.Open

	
	Set RSe = Server.CreateObject("ADODB.Recordset")
	SQLe = "SELECT * FROM TB_Entrevistas WHERE "&sql_matricula&sql_status_entrevista&"(DA_Entrevista BETWEEN #"&data_de&"# AND #"&data_ate&"#) order BY DA_Entrevista, HO_Entrevista"
	RSe.Open SQLe, CON4, 3, 3



check = 2
ordem_original=1
WHILE not RSe.EOF
	if check mod 2 =0 then
		cor = "tb_fundo_linha_par" 
	else 
		cor ="tb_fundo_linha_impar"
	end if 
  
	co_matric=RSe("CO_Matricula")
	da_entrevista=RSe("DA_Entrevista")
	ho_entrevista=RSe("HO_Entrevista")
	tp_entrevista=RSe("TP_Entrevista")
	no_participantes=RSe("NO_Participantes")
	st_entrevista=RSe("ST_Entrevista")
	co_agendado_com=RSe("CO_Agendado_com")
	tx_observaa=RSe("TX_Observa")
	co_usu_entrevista=RSe("CO_Usuario")
	
	if tp_entrevista="" or isnull(tp_entrevista) then
		tipo_entrevista=""
	else
	
		Set RST = Server.CreateObject("ADODB.Recordset")
		SQLT = "SELECT * FROM TB_Tipo_Entrevista Where tp_entrevista="&tp_entrevista
		RST.Open SQLT, CON0
	
		IF RST.EOF then
			tipo_entrevista=""	
		else
			tipo_entrevista=RST("TX_Descricao")
		end if	
	end if
				
	if co_agendado_com="" or isnull(co_agendado_com) then
		no_agendado=""
	else
		Set RSu = Server.CreateObject("ADODB.Recordset")
		SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& co_agendado_com
		RSu.Open SQLu, CON
	
		IF RSu.EOF then
			no_atendido =""	
		else
			no_atendido =RSu("NO_Usuario")
		end if
			
	end if
	
	'if co_usu_entrevista="" or isnull(co_usu_entrevista) then
	'	no_atendido=""
	'else
	'		Set RSu = Server.CreateObject("ADODB.Recordset")
	'		SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& co_usu_entrevista
	'		RSu.Open SQLu, CON
	'
	'	IF RSu.EOF then
	'		no_agendado=""	
	'	else
	'		no_agendado=RSu("NO_Usuario")
	'	end if
	'		
	'end if
	
	hora_split= Split(ho_entrevista,":")
	hora=hora_split(0)
	min=hora_split(1)
	
	ho_entrevista = hora&":"&min
	hora_entrevista = hora&":"&min

	Session("status_entrevista")=status_entrevista
	Session("data_de")=data_de
	Session("hora_de")=hora_de
	Session("data_inicio")=data_inicio
	Session("data_ate")=data_ate
	Session("hora_ate")=hora_ate
	Session("data_fim")=data_fim
	
	data_split= Split(da_entrevista,"/")
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
	da_show=dia&"/"&mes&"/"&ano&", "&hora&":"&min
	data_entrevista = dia&"/"&mes&"/"&ano
	
		Select case st_entrevista		
			case 1
			nome_status="Atendida"
			
			case 2
			nome_status="Cancelada"
			
			case 3
			nome_status="Pendente"	
		end select	
		
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& co_matric
	RS.Open SQL, CON1

	nome_aluno = RS("NO_Aluno")	
	
	Rs_ordena.AddNew		
	Rs_ordena.Fields("ordem_original").Value = ordem_original
	Rs_ordena.Fields("data").Value = da_show
	Rs_ordena.Fields("data_entrevista").Value = data_entrevista
	Rs_ordena.Fields("hora_entrevista").Value = hora_entrevista		
	Rs_ordena.Fields("co_matric").Value = co_matric
	Rs_ordena.Fields("nome").Value = nome_aluno
	Rs_ordena.Fields("tipo").Value = tipo_entrevista
	Rs_ordena.Fields("participantes").Value = no_participantes
	Rs_ordena.Fields("atendido_por").Value = no_atendido	
	Rs_ordena.Fields("nome_status").Value = nome_status
	
		check = check+1
		ordem_original = ordem_original +1
RSe.Movenext
'end if
WEND

'END IF

Select case ordem

case "dt"
Rs_ordena.Sort = "ordem_original ASC"

case "mt"
Rs_ordena.Sort = "co_matric ASC"

case "al"
Rs_ordena.Sort = "nome ASC"

case "en"
Rs_ordena.Sort = "tipo ASC"

case "pr"
Rs_ordena.Sort = "participantes ASC"

case "at"
Rs_ordena.Sort = "atendido_por ASC"

case "st"
Rs_ordena.Sort = "nome_status ASC"




end select

if Rs_ordena.EOF	then	
	desabilita = "S"
	sem_link="S"	
%>

                <tr> 
                  <td width="21">&nbsp;</td>
                  <td colspan="7" align="center" class="form_dado_texto">
                  Nenhuma entrevista cadastrada para os crit&eacute;rios informados</td>
                </tr>
<%else
	desabilita = "N"
	sem_link="N"	
			
    if cint(Request.QueryString("pagina"))<1 then
	intpagina = 1
	Rs_ordena.MoveFirst
    else
		if cint(Request.QueryString("pagina"))>Rs_ordena.PageCount then  
	    intpagina = Rs_ordena.PageCount
        else
    	intpagina = Request.QueryString("pagina")
		end if
    end if   


	
 Rs_ordena.PageSize = 30
 
if Request.QueryString("pagina")="" then
      intpagina = 1
	  Rs_ordena.MoveFirst
else
    if cint(Request.QueryString("pagina"))<1 then
	intpagina = 1
    else
		if cint(Request.QueryString("pagina"))>Rs_ordena.PageCount then  
	    intpagina = Rs_ordena.PageCount
        else
    	intpagina = Request.QueryString("pagina")
		end if
    end if   
 end if   

    Rs_ordena.AbsolutePage = intpagina
    intrec = 0
	check=2
	While intrec<Rs_ordena.PageSize and Not Rs_ordena.EoF
	
	 if check mod 2 =0 then
	  cor = "tb_fundo_linha_par" 
	 else cor ="tb_fundo_linha_impar"
	  end if
	  
da_show = Rs_ordena.Fields("data").Value
da_entrevista = Rs_ordena.Fields("data_entrevista").Value
ho_entrevista = Rs_ordena.Fields("hora_entrevista").Value
co_matric = Rs_ordena.Fields("co_matric").Value
nome_aluno = Rs_ordena.Fields("nome").Value
tipo_entrevista = Rs_ordena.Fields("tipo").Value
no_participantes = Rs_ordena.Fields("participantes").Value
no_atendido = Rs_ordena.Fields("atendido_por").Value
nome_status = Rs_ordena.Fields("nome_status").Value

	dados_entrevista=co_matric&"?"&da_entrevista&"?"&ho_entrevista
	  %>		
                <tr class="<%=cor%>"> 
                  <td width="21"> <input name="entrevista" type="checkbox" class="borda" id="entrevista" value="<%=dados_entrevista%>"></td>
                  <td width="90" align="center"> 
                    <a href="incluir.asp?ori=A&opt=<%response.Write(dados_entrevista)%>">
                    <%response.Write(da_show)%>
                    </a>
                    <div align="center"></div>
                    <div align="left"></div></td>
                  <td width="70" align="center"> 
                      <%response.Write(co_matric)%>
                      </td>
                  <td width="288"> 
                    <%response.Write(nome_aluno)%>
                    </td>
                  <td width="108" align="center"> 
                    <%response.Write(tipo_entrevista)%>
                    </td>
                  <td width="122"> <div align="center"> 
                      <%response.Write(no_participantes)%>
                    </div></td>
                  <td width="160"> <div align="center"> 
                      <%response.Write(no_atendido)%>
                    </div></td>
                  <td width="159" align="center"><%response.Write(nome_status)%></td>
                </tr>
                <%
		intrec=intrec+1
		check=check+1				
	Rs_ordena.movenext
	Wend
end if

if sem_link="N" AND Rs_ordena.PageCount > 1 then	 
%>
    <tr>
      <td colspan="8"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td class="tb_tit"><div align="center">                          </div></td>
          </tr>
          <tr> 
            <td class="tb_tit"><div align="center"> 
                <%
 
    if intpagina>1 then
    %>
        <a href="resumo.asp?pagina=<%=intpagina-1%>" class="linktres">Anterior</a> 
    <%
    end if 
	for contapagina=1 to Rs_ordena.PageCount 
		pagina=pagina*1
		IF contapagina=pagina then
			response.Write(contapagina)
		else
		%>
			<a href="resumo.asp?pagina=<%=contapagina%>" class="linktres"><%response.Write(contapagina)%></a> 
		<%
		end if
	next
    if StrComp(intpagina,Rs_ordena.PageCount)<>0 then  
    %>
        <a href="resumo.asp?pagina=<%=intpagina + 1%>" class="linktres">Próximo</a> 
    <%
    end if
	%>
</div></td>
      </tr>
    </table></td>
</tr>
<%end if	
 Rs_ordena.Close
Set Rs_ordena = Nothing

Set objPasta = Nothing
Set objFSO = Nothing
    %>
                         
                <tr class="<%=cor%>"> 
                  <td colspan="8">
                      <hr width="1000">
                    </td>
                </tr>
                <tr class="<%=cor%>"> 
                  <td colspan="8"><table width="1000" border="0" align="center" cellspacing="0">
                      <tr> 
                        <td width="20%"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                            <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','index.asp?nvg=WA-AL-MA-AGE');return document.MM_returnValue"value="Voltar">
                            </font></div></td>
                        <td width="20%" align="center"><%if trava<>"n" or permite_incluir = "N" then%>
                          <input name="Submit" type="submit" class="botao_prosseguir" id="Submit" disabled value="Marcar Nova">
                            <%else%>                        
                            <input name="Submit" type="submit" class="botao_prosseguir" id="Submit" onClick="MM_goToURL('parent','redireciona.asp?opt=<%=cod%>');return document.MM_returnValue" value="Marcar Nova">
                        <% end if%></td>
                        <td width="20%" align="center"><%if trava<>"n" or desabilita = "S" then%>
                            <input name="Submit" type="submit" class="botao_prosseguir" id="Submit" disabled value="Alterar">
                            <%else%>
                            <input name="Submit" type="submit" class="botao_prosseguir" id="Submit" value="Alterar">
                        <% end if%></td>
                        <td width="20%" align="center"><%if trava<>"n" or desabilita = "S" then%>
                            <input name="Submit" type="submit" class="botao_excluir" id="Submit" disabled value="Excluir">
                            <%else%>
                            <input name="Submit" type="submit" class="botao_excluir" id="Submit" value="Excluir">
                        <% end if%></td>
                        <td width="20%" align="center"><%if trava<>"n" or desabilita = "S"  then%>
                            <input name="Submit" type="submit" class="botao_prosseguir" id="Submit" disabled value="Conte&uacute;do">
                            <%else%>
                            <input name="Submit" type="submit" class="botao_prosseguir" id="Submit" value="Conte&uacute;do">
                        <% end if%></td>
                      </tr>
                  </table></td>
                </tr>
              </table>
            </form>
              </div></td>
        </tr>
      </table></td>
    </tr>

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