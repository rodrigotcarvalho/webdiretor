<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 
Session.LCID = 1046
nivel=4
ori = request.QueryString("ori")
opt= request.QueryString("res")
pagina=request.QueryString("pagina")

permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
trava=session("trava")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo


	cod= request.QueryString("cod_cons")

sessionobr=cod&"?"&ordem&"?"&status_entrevista&"?"&data_de&"?"&hora_de&"?"&data_inicio&"?"&data_ate&"?"&hora_ate&"?"&data_fim
trava=session("trava")
ocorr= request.form("ocorr")
session("obr")=sessionobr

aa = DatePart("yyyy", now)
mm = DatePart("m", now) 
dd = DatePart("d", now) 
	

dados_msg  = cod




		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON5 = Server.CreateObject("ADODB.Connection") 
		ABRIR5 = "DBQ="& CAMINHO_ei & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON5.Open ABRIR5			
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0	

	
	permite_incluir = "S"	
	

		
		





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
'end if


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
      <%call mensagens(nivel,663,2,ori) %>
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
      <%call mensagens(nivel,645,0,"R19") %>
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
        </tr>
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
              </table>
            </td>
        </tr>
        <tr> 
          <td height="10" colspan="2" class="tb_tit"
>Entrevista Inicial</td>
        </tr>
       <tr > 
          <td height="154" colspan="2" valign="top"> 
		  
 <form action="bd.asp" method="post" name="busca" id="busca">		  
		  
              <table width="1000" border="0" cellspacing="0" cellpadding="0">
                <%			

	
'	Set RSe = Server.CreateObject("ADODB.Recordset")
'	SQLe = "SELECT * FROM TB_Entrevistas WHERE CO_Matricula = "&cod&" AND TP_Entrevista = 1"
'	RSe.Open SQLe, CON4
'
'check = 2
'ordem_original=1
'
'if RSe.EOF	then	
'	desabilita = "S"
'	sem_link="S"	
%>
<!--
                <tr> 
                  <td height="20" colspan="8" align="center" class="form_dado_texto">
                  Entrevista inicial não cadastrada para o aluno informados</td>
                </tr>
--><%'else
'
'	if check mod 2 =0 then
'		cor = "tb_fundo_linha_par" 
'	else 
'		cor ="tb_fundo_linha_impar"
'	end if 
'  
'
'	co_matric=RSe("CO_Matricula")
'	da_entrevista=RSe("DA_Entrevista")
'	ho_entrevista=RSe("HO_Entrevista")
'	tp_entrevista=RSe("TP_Entrevista")
'	no_participantes=RSe("NO_Participantes")
'	st_entrevista=RSe("ST_Entrevista")
'	co_agendado_com=RSe("CO_Agendado_com")
'	tx_observaa=RSe("TX_Observa")
'	co_usu_entrevista=RSe("CO_Usuario")
'	
'	if tp_entrevista="" or isnull(tp_entrevista) then
'		tipo_entrevista=""
'	else
'	
'		Set RST = Server.CreateObject("ADODB.Recordset")
'		SQLT = "SELECT * FROM TB_Tipo_Entrevista Where tp_entrevista="&tp_entrevista
'		RST.Open SQLT, CON0
'	
'		IF RST.EOF then
'			tipo_entrevista=""	
'		else
'			tipo_entrevista=RST("TX_Descricao")
'		end if	
'	end if
'				
'	if co_agendado_com="" or isnull(co_agendado_com) then
'		no_agendado=""
'	else
'		Set RSu = Server.CreateObject("ADODB.Recordset")
'		SQLu = "SELECT * FROM TB_Usuario WHERE CO_Usuario ="& co_agendado_com
'		RSu.Open SQLu, CON
'	
'		IF RSu.EOF then
'			no_atendido =""	
'		else
'			no_atendido =RSu("NO_Usuario")
'		end if
'			
'	end if

	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Entrevistas_Inicial WHERE CO_Matricula = "&cod
	RS.Open SQL, CON5

check = 2
ordem_original=1

if RS.EOF	then

	dat_adapta = dd&"/"&mm&"/"&aa
	irmao1 = "" 
	idade1 = "" 
	irmao2 = "" 
	idade2 = ""
	irmao3 = ""
	idade3 = "" 
	outros = ""
	desejada = ""
	esperada = ""
	como_passou = ""
	normal = ""
	termo = ""
	prematuro = ""
	cesariana = ""
	dia_nascimento = ""
	materna = ""
	pegou_bem = ""
	artificial = "" 
	adaptacao_mudanca = "" 
	chupava_dedo = ""
	chupeta = ""
	alimentacao = ""	
	dificuldade_alimentacao = ""
	sentou = ""
	arrastou = "" 
	engatinhou = ""
	andou = ""
	linguagem = ""
	dificuldade_fala = "" 
	pedalar = ""
	infeccoes = "" 
	alergias = ""
	outras_infeccoes = ""
	antitermico = ""
	antecedentes = ""
	divertimentos = ""
	higiene = ""
	controle = "" 
	sono = ""
	gosta_fazer = ""
	caracteristicas = ""



else	

	dat_adapta = RS("DA_Adapta") 
	irmao1 = RS("NO_Irmao1") 
	idade1 = RS("ID_Irmao1") 
	irmao2 = RS("NO_Irmao2") 
	idade2 = RS("ID_Irmao2") 
	irmao3 = RS("NO_Irmao3") 
	idade3 = RS("ID_Irmao3") 
	outros = RS("TX_Outras_Pessoas") 
	desejada = RS("TX_ISC_Desejada") 
	esperada = RS("TX_ISC_Esperada") 
	como_passou = RS("TX_ISC_Como_grav") 
	normal = RS("TX_ISC_Normal") 
	termo = RS("TX_ISC_Termo") 
	prematuro = RS("TX_ISC_Prema") 	
	cesariana = RS("TX_ISC_Cesariana") 
	dia_nascimento = RS("TX_ISC_Como_Parto") 
	materna = RS("TX_ISC_Materna") 
	pegou_bem = RS("TX_ISC_Pegou") 
	artificial = RS("TX_ISC_Artificial") 
	adaptacao_mudanca = RS("TX_ISC_Como_mud") 
	chupava_dedo = RS("TX_ISC_chupava") 
	chupeta = RS("TX_ISC_chupeta") 
	alimentacao = RS("TX_ISC_alim") 
	dificuldade_alimentacao = RS("TX_ISC_Como_alim") 	
	sentou = RS("TX_DP_Sentou") 
	arrastou = RS("TX_DP_Arrastou") 
	engatinhou = RS("TX_DP_Enga") 
	andou = RS("TX_DP_Andou") 
	linguagem = RS("TX_DP_Ling") 
	dificuldade_fala = RS("TX_DP_Obs") 
	pedalar = RS("TX_DP_Anda_bem") 
	infeccoes = RS("TX_AP_Infec") 
	alergias = RS("TX_AP_alergia") 
	outras_infeccoes = RS("TX_AP_outros") 
	antitermico = RS("TX_AP_Antit") 
	antecedentes = RS("TX_AP_Antece") 
	divertimentos = RS("TX_DF") 
	higiene = RS("TX_AH_Hig") 
	controle = RS("TX_AH_Como") 
	sono = RS("TX_AH_Sono") 
	gosta_fazer = RS("TX_IN_Sob") 
	caracteristicas = RS("TX_IN_Carac") 
	'co_user_bd = RS("CO_Usuario")

end if

dt_adpt = split(dat_adapta, "/")
dd = dt_adpt(0)
mm = dt_adpt(1)
aa = dt_adpt(2)
%>                
<tr>
  <td colspan="2" align="center" class="form_dado_texto"><table width="100%" border="0" cellspacing="0" cellpadding="0">
    <tr class="form_dado_texto">
      <td width="136" align="right">Data de Adapta&ccedil;&atilde;o:&nbsp;</td>
      <td width="864"> <select name="dia_de" id="dia_de" class="select_style" <%response.Write(disable)%>>
                          <%

	for dia= 1 to 31
dd=dd*1
dia=dia*1	
		if dd=dia then
			dia_selected = "selected"
		else
			dia_selected = ""	
		end if	
		
		if dia<10 then
			dia_txt="0"&dia
		else	
			dia_txt=dia		
		end if		
	
%>
                          <option value="<%response.Write(dia)%>" <%response.Write(dia_selected)%>>
                            <%response.Write(dia_txt)%>
                            </option>
                          <%next%>
                        </select>
                        /
                        <select name="mes_de" id="mes_de" class="select_style" <%response.Write(disable)%>>
                          <%

	for mes= 1 to 12
		mm=mm*1
		mes=mes*1	
		if mm=mes then
			mes_selected = "selected"
		else
			mes_selected = ""	
		end if	
		
		Select case mes
		
			case 1
			mes_txt="janeiro"
			
			case 2
			mes_txt="fevereiro"
			
			case 3
			mes_txt="mar&ccedil;o"
			
			case 4
			mes_txt="abril"
			
			case 5
			mes_txt="maio"
	
			case 6
			mes_txt="junho"
			
			case 7
			mes_txt="julho"
			
			case 8
			mes_txt="agosto"		
			
			case 9
			mes_txt="setembro"
	
			case 10
			mes_txt="outubro"
			
			case 11
			mes_txt="novembro"
			
			case 12
			mes_txt="dezembro"				
		end select			
	
%>
                          <option value="<%response.Write(mes)%>" <%response.Write(mes_selected)%>>
                            <%response.Write(mes_txt)%>
                            </option>
                          <%next%>
                        </select>
                        /
                        <select name="ano_de" class="select_style" id="ano_de" <%response.Write(disable)%>>
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
		ano_bd = ano_bd+1	
				%>
                          <option value="<%=ano_bd%>"><%=ano_bd%></option>                
                        </select><input name="cod_cons" id="cod_cons" type="hidden" value="<%response.write(cod)%>"></td>
    </tr>
  </table></td>
</tr>
<tr> 
                  <td colspan="2" align="center" class="tb_subtit">
                      Irm&atilde;os</td>
                </tr>                                      <tr> 
                  <td height="20" colspan="2" align="center" valign="top" class="form_dado_texto">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">

  <tr>
    <td width="100" align="right"><span class="form_dado_texto">Nome:&nbsp;</span></td>
    <td width="400"><input name="irmao1" type="text" class="form_dado_texto" id="irmao1" value="<%response.write(irmao1)%>" size="50" maxlength="50"></td>
    <td width="100" align="right"><span class="form_dado_texto">Idade:&nbsp;</span></td>
    <td width="400"><input name="idade1" type="text" class="form_dado_texto" id="idade1" value="<%response.write(idade1)%>" size="3" maxlength="3"></td>
  </tr>
  <tr>
    <td align="right"><span class="form_dado_texto">Nome:&nbsp;</span></td>
    <td><input name="irmao2" type="text" class="form_dado_texto" id="irmao2" value="<%response.write(irmao2)%>" size="50" maxlength="50"></td>
    <td align="right"><span class="form_dado_texto">Idade:&nbsp;</span></td>
    <td><input name="idade2" type="text" class="form_dado_texto" id="idade2" value="<%response.write(idade2)%>" size="3" maxlength="3"></td>
  </tr>
  <tr>
    <td align="right"><span class="form_dado_texto">Nome:&nbsp;</span></td>
    <td><input name="irmao3" type="text" class="form_dado_texto" id="irmao3" value="<%response.write(irmao3)%>" size="50" maxlength="50"></td>
    <td align="right"><span class="form_dado_texto">Idade:&nbsp;</span></td>
    <td><input name="idade3" type="text" class="form_dado_texto" id="idade3" value="<%response.write(idade3)%>" size="3" maxlength="3"></td>
  </tr>
</table>
</td>
                </tr>  
                <tr class="tb_tit"> 
                  <td colspan="2" align="center" class="tb_subtit">
                     Outras pessoas residindo com a fam&iacute;lia</td>
                </tr>                                      <tr> 
                  <td height="20" colspan="2" align="center" valign="top" class="form_dado_texto">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">

  <tr>
    <td width="100" align="right" valign="top"><span class="form_dado_texto">Nomes:&nbsp;</span></td>
    <td><textarea name="outros" cols="170" rows="5" id="outros" class="form_dado_texto"><%response.write(outros)%></textarea></td>
    </tr>
</table>
</td>
                </tr>  
                <tr class="tb_tit"> 
                  <td colspan="2" align="center" class="tb_subtit">
                     Informa&ccedil;&otilde;es sobre a crian&ccedil;a</td>
                </tr>                                      <tr> 
                  <td height="20" colspan="2" align="center" valign="top" class="form_dado_texto">
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">

  <tr>
    <td width="99" align="right"><span class="form_dado_texto">Gravidez:&nbsp;</span></td>
    <td width="140" class="form_dado_texto">Desejada</td>
    <td width="732"><input name="desejada" type="text" class="form_dado_texto" id="desejada" value="<%response.write(desejada)%>" size="140"></td>
    </tr>
  <tr>
    <td width="99" align="right"><span class="form_dado_texto">&nbsp;</span></td>
    <td class="form_dado_texto">Esperada</td>
    <td><input name="esperada" type="text" class="form_dado_texto" id="esperada" value="<%response.write(esperada)%>" size="140"></td>
    </tr>
  <tr>
    <td width="99" align="right">&nbsp;</td>
    <td align="right" class="form_dado_texto">Como Passou:&nbsp;</td>
    <td><input name="como_passou" type="text" class="form_dado_texto" id="como_passou" value="<%response.write(como_passou)%>" size="140" maxlength="140"></td>
    </tr>
  <tr>
    <td width="99" align="right">&nbsp;</td>
    <td class="form_dado_texto">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="99" align="right"><span class="form_dado_texto">Parto:&nbsp;</span></td>
    <td width="140" class="form_dado_texto">A Termo</td>
    <td><input name="termo" type="text" class="form_dado_texto" id="termo" value="<%response.write(termo)%>" size="140" maxlength="140"></td>
  </tr>
  <tr>
    <td width="99" align="right"><span class="form_dado_texto">&nbsp;</span></td>
    <td width="140" class="form_dado_texto">Prematuro</td>
    <td><input name="prematuro" type="text" class="form_dado_texto" id="prematuro" value="<%response.write(prematuro)%>" size="140" maxlength="140"></td>
  </tr>
  <tr>
    <td width="99" align="right">&nbsp;</td>
    <td width="140" class="form_dado_texto">Normal</td>
    <td><input name="normal" type="text" class="form_dado_texto" id="normal" value="<%response.write(normal)%>" size="140" maxlength="140"></td>
  </tr>
  <tr>
    <td width="99" align="right">&nbsp;</td>
    <td width="140" class="form_dado_texto">Cesariana</td>
    <td><input name="cesariana" type="text" class="form_dado_texto" id="cesariana" value="<%response.write(cesariana)%>" size="140" maxlength="140"></td>
  </tr>
  <tr>
    <td width="99" align="right">&nbsp;</td>
    <td colspan="2" class="form_dado_texto"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr class="form_dado_texto">
        <td width="190" height="20" align="right">Como foi o dia do nascimento?&nbsp;</td>
        <td width="78%" rowspan="2" valign="top"><textarea name="dia_nascimento" cols="120" rows="3" class="form_dado_texto" id="dia_nascimento"><%response.write(dia_nascimento)%></textarea></td>
        </tr>
      <tr class="form_dado_texto">
        <td width="190" height="20" align="right">&nbsp;</td>
        </tr>
    </table></td>
    </tr>
  <tr>
    <td align="right">&nbsp;</td>
    <td class="form_dado_texto">&nbsp;</td>
    <td>&nbsp;</td>
  </tr>
  <tr>
    <td width="99" align="right"><span class="form_dado_texto">Alimenta&ccedil;&atilde;o:&nbsp;</span></td>
    <td class="form_dado_texto">Materna</td>
    <td><input name="materna" type="text" class="form_dado_texto" id="materna" value="<%response.write(materna)%>" size="140" maxlength="140"></td>
  </tr>
  <tr>
    <td width="99" align="right">&nbsp;</td>
    <td align="right" class="form_dado_texto">Pegou Bem o Seio?&nbsp;</td>
    <td><input name="pegou_bem" type="text" class="form_dado_texto" id="pegou_bem" value="<%response.write(pegou_bem)%>" size="140" maxlength="140"></td>
  </tr>
  <tr>
    <td width="99" align="right">&nbsp;</td>
    <td class="form_dado_texto">Artificial</td>
    <td><input name="artificial" type="text" class="form_dado_texto" id="artificial" value="<%response.write(artificial)%>" size="140" maxlength="140"></td>
  </tr>
  <tr>
    <td width="99" align="right">&nbsp;</td>
    <td colspan="2" class="form_dado_texto"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr class="form_dado_texto">
        <td width="190" height="20" align="right">Como a aceita&ccedil;&atilde;o da Mudan&ccedil;a?&nbsp;</td>
        <td width="78%" rowspan="2" valign="top"><textarea name="adaptacao_mudanca" cols="120" rows="3" class="form_dado_texto" id="adaptacao_mudanca"><%response.write(adaptacao_mudanca)%></textarea></td>
        </tr>
      <tr class="form_dado_texto">
        <td width="190" height="20" align="right">&nbsp;</td>
        </tr>
    </table></td>
    </tr>
  <tr>
    <td width="99" align="right">&nbsp;</td>
    <td class="form_dado_texto">Chupava dedo:</td>
    <td><input name="chupava_dedo" type="text" class="form_dado_texto" id="chupava_dedo" value="<%response.write(chupava_dedo)%>" size="140" maxlength="140"></td>
  </tr>
  <tr>
    <td width="99" align="right">&nbsp;</td>
    <td class="form_dado_texto">Chupeta:</td>
    <td><input name="chupeta" type="text" class="form_dado_texto" id="chupeta" value="<%response.write(chupeta)%>" size="140" maxlength="140"></td>
  </tr>
  <tr>
    <td width="99" align="right">&nbsp;</td>
    <td class="form_dado_texto">Alimenta&ccedil;&atilde;o:</td>
    <td><input name="alimentacao" type="text" class="form_dado_texto" id="alimentacao" value="<%response.write(alimentacao)%>" size="140" maxlength="140"></td>
  </tr>
  <tr>
    <td width="99" align="right">&nbsp;</td>
    <td colspan="2" class="form_dado_texto"><table width="100%" border="0" cellspacing="0" cellpadding="0">
      <tr class="form_dado_texto">
        <td height="20" align="left">Como a fam&iacute;lia reage quando h&aacute; dificuldade na alimenta&ccedil;&atilde;o?&nbsp;</td>
        </tr>
      <tr class="form_dado_texto">
        <td height="20" align="left"><textarea name="dificuldade_alimentacao" cols="170" rows="3" class="form_dado_texto" id="dificuldade_alimentacao"><%response.write(dificuldade_alimentacao)%></textarea></td>
        </tr>
    </table></td>
    </tr>
                  </table>
</td>
                </tr>
                <tr>
                  <td height="20" colspan="2" align="center" class="tb_subtit">Desenvolvimento Psicomotor</td>
                </tr>
                <tr>
                  <td height="20" colspan="2" align="center" valign="top" class="form_dado_texto"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="100" align="right">&nbsp;</td>
                      <td width="174" class="form_dado_texto">Sentou:</td>
                      <td width="726"><input name="sentou" type="text" class="form_dado_texto" id="sentou" value="<%response.write(sentou)%>" size="140" maxlength="140"></td>
                    </tr>
                    <tr>
                      <td align="right">&nbsp;</td>
                      <td class="form_dado_texto">Arrastou:</td>
                      <td><input name="arrastou" type="text" class="form_dado_texto" id="arrastou" value="<%response.write(arrastou)%>" size="140" maxlength="140"></td>
                    </tr>
                    <tr>
                      <td align="right">&nbsp;</td>
                      <td class="form_dado_texto">Engatinhou:</td>
                      <td><input name="engatinhou" type="text" class="form_dado_texto" id="engatinhou" value="<%response.write(engatinhou)%>" size="140" maxlength="140"></td>
                    </tr>
                    <tr>
                      <td align="right">&nbsp;</td>
                      <td class="form_dado_texto">Andou:</td>
                      <td><input name="andou" type="text" class="form_dado_texto" id="andou" value="<%response.write(andou)%>" size="140" maxlength="140"></td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                      <td class="form_dado_texto">Linguagem:</td>
                      <td><input name="linguagem" type="text" class="form_dado_texto" id="linguagem" value="<%response.write(linguagem)%>" size="140" maxlength="140"></td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                      <td colspan="2" class="form_dado_texto"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr class="form_dado_texto">
                          <td height="20" align="left">Voc&ecirc;s observam alguma dificuldade na fala? Qual?&nbsp;</td>
                        </tr>
                        <tr class="form_dado_texto">
                          <td height="20" align="left"><textarea name="dificuldade_fala" cols="175" rows="3" class="form_dado_texto" id="dificuldade_fala"><%response.write(dificuldade_fala)%></textarea></td>
                        </tr>
                      </table></td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                      <td colspan="2" class="form_dado_texto"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr class="form_dado_texto">
                          <td height="20" align="left">Anda bem em brinquedos que precise pedalar?&nbsp;</td>
                        </tr>
                        <tr class="form_dado_texto">
                          <td height="20" align="left"><input name="pedalar" type="text" class="form_dado_texto" id="pedalar" value="<%response.write(pedalar)%>" size="175" maxlength="175"></td>
                        </tr>
                      </table></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="20" colspan="2" align="center" class="tb_subtit">Antecedentes Patol&oacute;gicos</td>
                </tr>
                <tr>
                  <td height="20" colspan="2" align="center" class="form_dado_texto"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="100" align="right">&nbsp;</td>
                      <td width="174" class="form_dado_texto">Infec&ccedil;&otilde;es:</td>
                      <td width="726"><input name="infeccoes" type="text" class="form_dado_texto" id="infeccoes" value="<%response.write(infeccoes)%>" size="140" maxlength="140"></td>
                    </tr>
                    <tr>
                      <td align="right">&nbsp;</td>
                      <td class="form_dado_texto">Alergias:</td>
                      <td><input name="alergias" type="text" class="form_dado_texto" id="alergias" value="<%response.write(alergias)%>" size="140" maxlength="140"></td>
                    </tr>
                    <tr>
                      <td align="right">&nbsp;</td>
                      <td class="form_dado_texto">Outros:</td>
                      <td><input name="outras_infeccoes" type="text" class="form_dado_texto" id="outras_infeccoes" value="<%response.write(outras_infeccoes)%>" size="140" maxlength="140"></td>
                    </tr>
                    <tr>
                      <td align="right">&nbsp;</td>
                      <td class="form_dado_texto">Antit&eacute;rmico:</td>
                      <td><input name="antitermico" type="text" class="form_dado_texto" id="antitermico" value="<%response.write(antitermico)%>" size="140" maxlength="140"></td>
                    </tr>
                    <tr>
                      <td>&nbsp;</td>
                      <td colspan="2" class="form_dado_texto"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr class="form_dado_texto">
                          <td height="20" align="left">Antecedentes familiares</td>
                          </tr>
                        <tr class="form_dado_texto">
                          <td height="20" align="left"><textarea name="antecedentes" cols="175" rows="3" class="form_dado_texto" id="antecedentes"><%response.write(antecedentes)%></textarea></td>
                          </tr>
                        </table></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="20" colspan="2" align="center" class="tb_subtit">Divertimentos da Fam&iacute;lia</td>
                </tr>
                <tr>
                  <td width="100" height="20" align="center" valign="top" class="form_dado_texto">&nbsp;</td>
                  <td height="20" align="left" valign="top" class="form_dado_texto"><textarea name="divertimentos" cols="175" rows="3" class="form_dado_texto" id="divertimentos"><%response.write(divertimentos)%></textarea></td>
                </tr>
                <tr>
                  <td height="20" colspan="2" align="center" class="tb_subtit">Aquisi&ccedil;&atilde;o de h&aacute;bitos</td>
                </tr>
                <tr>
                  <td height="20" colspan="2" align="center" valign="top" class="form_dado_texto"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="136" align="right" class="form_dado_texto">Higiene (banho etc):&nbsp;</td>
                      <td width="864" class="form_dado_texto"><input name="higiene" type="text" class="form_dado_texto" id="higiene" value="<%response.write(higiene)%>" size="170" maxlength="170"></td>
                      </tr>
                    <tr align="right">
                      <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr class="form_dado_texto">
                          <td height="20" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Como e quando foi treinado (controle dos esfincteres)?</td>
                        </tr>
                        <tr class="form_dado_texto">
                          <td height="20" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input name="controle" type="text" class="form_dado_texto" id="controle" value="<%response.write(controle)%>" size="190" maxlength="190"></td>
                        </tr>
                      </table></td>
                      </tr>
                    <tr>
                      <td width="136" height="18" align="right" class="form_dado_texto">Sono:&nbsp;</td>
                      <td class="form_dado_texto"><input name="sono" type="text" class="form_dado_texto" id="sono" value="<%response.write(sono)%>" size="170" maxlength="170"></td>
                      </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="20" colspan="2" align="center" class="tb_subtit">Interesses</td>
                </tr>
                <tr>
                  <td height="20" colspan="2" align="center" valign="top" class="form_dado_texto"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr class="form_dado_texto">
                      <td height="20" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Sob o ponto de vista de voc&ecirc;s, o que a crian&ccedil;a gosta mais de fazer? Tem alguma coisa que n&atilde;o goste?</td>
                    </tr>
                    <tr class="form_dado_texto">
                      <td height="20" align="left">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                        <textarea name="gosta_fazer" cols="190" rows="3" class="form_dado_texto" id="gosta_fazer"><%response.write(gosta_fazer)%>
                        </textarea></td>
                    </tr>
                  </table></td>
                </tr>
                <tr>
                  <td height="20" colspan="2" align="center" class="tb_subtit">Caracter&iacute;sticas</td>
                </tr>
                <tr>
                  <td height="20" colspan="2" align="left" class="form_dado_texto">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Voc&ecirc;s consideram seu(ua) filho(a) uma crian&ccedil;a f&aacute;cil de lidar? Como reage quando &eacute; contrariada? Quem em casa d&aacute; mais aten&ccedil;&atilde;o &agrave; crian&ccedil;a? Escolaridade anterior?</td>
                </tr>
                <tr>
                  <td height="20" colspan="2" align="left" class="form_dado_texto">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
                  <textarea name="caracteristicas" cols="190" rows="3" class="form_dado_texto" id="caracteristicas"><%response.write(caracteristicas)%>
                    </textarea></td>
                </tr>
                <tr>
                  <td height="20" colspan="2" align="center" valign="top" class="form_dado_texto">&nbsp;</td>
                </tr>
                <tr>
                  <td height="20" colspan="2" align="center" valign="top" class="form_dado_texto">&nbsp;</td>
                </tr>
                <tr>
                  <td height="20" colspan="2" align="center" valign="top" class="form_dado_texto">&nbsp;</td>
                </tr>                  
                
                         <%
		intrec=intrec+1
		check=check+1				

'end if



    %>
                         
                <tr class="<%=cor%>"> 
                  <td colspan="2">
                      <hr width="1000">
                  </td>
                </tr>
                <tr class="<%=cor%>"> 
                  <td colspan="2"><table width="1000" border="0" align="center" cellspacing="0">
                    <tr>
                      <td width="33%"><div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">
                        <input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','index.asp?nvg=WA-AL-MA-AGE');return document.MM_returnValue"value="Voltar">
                      </font></div></td>
                      <td width="33%" align="center">&nbsp;</td>
                      <td width="33%" align="center"><%if trava<>"n" or desabilita = "S"  then%>
                        <input name="Submit4" type="submit" class="botao_prosseguir" id="Submit3" disabled value="Salvar">
                        <%else%>
                        <input name="Submit4" type="submit" class="botao_prosseguir" id="Submit3" value="Salvar">
                        <% end if%></td>
                    </tr>
                  </table>                   </td>
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