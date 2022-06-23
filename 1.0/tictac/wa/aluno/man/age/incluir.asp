<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<% 

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")

chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo

trava=session("trava")

ori = request.QueryString("ori")
entrevista= request.querystring("opt")
opt = request.querystring("res")

dados_entrevista = split(entrevista,"?")
cod=dados_entrevista(0)
if ori<>"I" then
	da_entrevista = dados_entrevista(1)
	ho_entrevista = dados_entrevista(2)
	
	hora_split= Split(ho_entrevista,":")
	hora=hora_split(0)
	min=hora_split(1)
	
	ho_entrevista=hora&":"&min
	
	obr=cod&"?"&da_entrevista&"?"&ho_entrevista
	
	data_split= Split(da_entrevista,"/")
	dia_e=data_split(0)
	mes_e=data_split(1)
	ano_e=data_split(2)
	
	da_entrevista_cons=mes_e&"/"&dia_e&"/"&ano_e	
	ho_entrevista_cons=hora&":"&min
	
	data_exibe	= dia_e&"/"&mes_e&"/"&ano_e		

'	dia=dia*1
'	mes=mes*1
'	hora=hora*1
'	min=min*1
'	
'	if dia<10 then
'		dia="0"&dia
'	end if
'	if mes<10 then
'		mes="0"&mes
'	end if
'	if hora<10 then
'		hora="0"&hora
'	end if
'	if min<10 then
'		min="0"&min
'	end if	
	dd=dia_e
	mm=mes_e
	aa=ano_e
	he=hora
	mn=min	
else
	aa = DatePart("yyyy", now)
	mm = DatePart("m", now) 
	dd = DatePart("d", now) 
	he = DatePart("h", now) 
	mn = DatePart("n", now) 	
	tp_entrevista = 2
end if	

obr=session("obr")
session("obr")=obr

	dados_msg=cod&"$!$"&da_entrevista&"$!$"&ho_entrevista

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
				
		Set CONp = Server.CreateObject("ADODB.Connection") 
		ABRIRp = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONp.Open ABRIRp		
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		
		
if cod="" or isnull(cod) then		
else
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
	RS.Open SQL, CON1
	
	codigo = RS("CO_Matricula")
	nome_prof = RS("NO_Aluno")

	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& cod
	RS.Open SQL, CON1


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
end if

if ori <> "I" then
		Set RSo = Server.CreateObject("ADODB.Recordset")
		SQLo = "SELECT * FROM TB_Entrevistas WHERE CO_Matricula ="& cod&" AND (DA_Entrevista=#"&da_entrevista_cons&"# AND mid(HO_Entrevista,1,16)=#12/30/1899 "&ho_entrevista_cons&"#)" 
		RSo.Open SQLo, CON4
		
	da_entrevista=RSo("DA_entrevista")
	ho_entrevista=RSo("HO_Entrevista")
	tp_entrevista=RSo("TP_Entrevista")
	partic_entrevista=RSo("NO_Participantes")
	st_entrevista=RSo("ST_Entrevista")
	ag_entrevista=RSo("CO_Agendado_com")
	ob_entrevista=RSo("TX_Observa")
	cu_entrevista=RSo("CO_Usuario")		
		
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
<%else %><body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"> 
<%end if %>
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
<%
 if opt="ok1" then    %>    
            <tr>    
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,319,2,0) %>
    </td>
			  </tr>
	<% end if 
if opt="ok2" then    %>    
            <tr>    
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,324,2,0) %>
    </td>
			  </tr>		  		  
<% end if	%>          
<%if ori="I" then
	action = "bd.asp?opt=inc"
%> 
	 <tr>              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,321,0,0) %>
    </td>
	</tr>	    
<%elseif ori="A" then
	action = "bd.asp?opt=alt"
%> 
	 <tr>     
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,322,0,0) %>
    </td>
	</tr>	    
<%elseif ori="C" then
	action = "bd.asp?opt=con"
	 ' if tp_entrevista > 1 then	
%>  
	 <tr>        
    <td height="10" colspan="5" valign="top"> 
      <%

	 	call mensagens(nivel,645,0,"R18") 
	%>
    </td>  
	</tr>
<%	  'end if%>    
	 <tr>        
    <td height="10" colspan="5" valign="top"> 
      <%
	  'if tp_entrevista > 1 then
	 	call mensagens(nivel,323,0,0) 
	  'else
	 '	call mensagens(nivel,325,1,0) 	  
	  'end if	%>
    </td>  
	</tr>		     
<%end if%> 

<%	st_entrevista=st_entrevista*1
	if st_entrevista = 1 then
		disable = "disabled"
   else
   		disable = ""
   end if

%>	          
	  
        <form action="<%response.Write(action)%>" method="post" name="busca" id="busca" onSubmit="return checksubmit()">
          <tr>
      <td valign="top"> 
        <table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
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
                    <input name="cod" type="hidden" value="<%=codigo%>">
                    <%response.Write(codigo)%>
                    </font></td>
                  <td width="6%" height="10"> <div align="right"><font class="form_dado_texto"> 
                      Nome: </font></div></td>
                  <td width="66%" height="10"><font class="form_dado_texto"> 
                    <%response.Write(nome_prof)%>
                    <input name="nome" type="hidden" class="select_style" id="nome" value="<%response.Write(nome_prof)%>" size="75" maxlength="50">
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
            <td colspan="2"><table width="100%" border="0" cellspacing="0">
                <tr class="tb_subtit"> 
                  <td width="100" height="10"> <div align="center"> 
                  Ano</div></td>
                  <td width="100" height="10"> <div align="center">Matr&iacute;cula</div></td>
                  <td width="100" height="10"> <div align="center">Cancelamento</div></td>
                  <td width="100" height="10"> <div align="center"> Situa&ccedil;&atilde;o</div></td>
                  <td width="150" height="10"> <div align="center">Unidade</div></td>
                  <td width="150" height="10"> <div align="center">Curso</div></td>
                  <td width="221" height="10"> <div align="center"> Etapa</div></td>
                  <td height="10"> 
                    <div align="center">Turma </div></td>
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
                  <td width="221" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_etapa)%>
                  </font></div></td>
                  <td height="10"> 
                    <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                  </font></div> 
                  <div align="center"></div></td>
                </tr>
                <tr class="tb_corpo">
                  <td height="10">&nbsp;</td>
                  <td height="10">&nbsp;</td>
                  <td height="10">&nbsp;</td>
                  <td height="10">&nbsp;</td>
                  <td height="10">&nbsp;</td>
                  <td height="10">&nbsp;</td>
                  <td height="10">&nbsp;</td>
                  <td height="10">&nbsp;</td>
                </tr>
                <tr>
                  <td height="10" colspan="8"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr class="tb_tit">
                      <td height="10" colspan="7">Entrevista</td>
                      </tr>
                    <tr class="tb_subtit">
                      <td height="10" align="center">Data </td>
                      <td height="10" align="center">Hora</td>
                      <td height="10" align="center">Tipo</td>
                      <%if ori="I" then  %>
                      <td height="10" colspan="3" align="center">Atendimento</td>  
                      <%else%>                    
                      <td height="10" colspan="2" align="center">Atendimento</td>
                      <td align="center">Status</td>
                      <%end if%>
                    </tr>
                    <tr class="form_dado_texto">
                      <td height="10" align="center"><% if ori = "C" then
		response.Write(data_exibe)%>
                        <input name="dia_de" type="hidden" value="<%response.Write(dd)%>">
                        <input name="mes_de" type="hidden" value="<%response.Write(mm)%>">
                        <input name="ano_de" type="hidden" value="<%response.Write(aa)%>">
                        <%

 else %>
                        <input name="dia_original" type="hidden" value="<%response.Write(dd)%>">
                        <input name="mes_original" type="hidden" value="<%response.Write(mm)%>">
                        <input name="ano_original" type="hidden" value="<%response.Write(aa)%>"> 
                        <input name="dia_de_disable" type="hidden" value="<%response.Write(dd)%>">
                        <input name="mes_de_disable" type="hidden" value="<%response.Write(mm)%>">
                        <input name="ano_de_disable" type="hidden" value="<%response.Write(aa)%>">                        
                        <select name="dia_de" id="dia_de" class="select_style" <%response.Write(disable)%>>
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
		
			aa=aa*1
			ano_bd=ano_bd*1

				if aa=ano_bd then%>
                          <option value="<%=ano_bd%>" selected><%=ano_bd%></option>
                          <%else%>
                          <option value="<%=ano_bd%>"><%=ano_bd%></option>
                          <%end if
		RS0.MOVENEXT
		WEND 	
		ano_bd = ano_bd+1	
				if aa=ano_bd then%>
                          <option value="<%=ano_bd%>" selected><%=ano_bd%></option>
                          <%else%>
                          <option value="<%=ano_bd%>"><%=ano_bd%></option>
                          <%end if
%>						             
                        </select>
                        <%end if%></td>
                      <td height="10" align="center"><% if ori = "C" then
		response.Write(ho_entrevista_cons) %>
                        <input name="hora_de" type="hidden" value="<%response.Write(he)%>">
                        <input name="min_de" type="hidden" value="<%response.Write(mn)%>">
                        <% 
else %>                 <input name="hora_original" type="hidden" value="<%response.Write(he)%>">
                        <input name="min_original" type="hidden" value="<%response.Write(mn)%>">
                        <input name="hora_de_disable" type="hidden" value="<%response.Write(he)%>">
                        <input name="min_de_disable" type="hidden" value="<%response.Write(mn)%>">                        
                        <select name="hora_de" id="hora_de" class="select_style" <%response.Write(disable)%>>
                          <% 
							 For n = 0 to 23
								 he=he*1
								 n=n*1
								if n<10 then
									h_exibe="0"&n
								else
									h_exibe=n									
								end if	
								 he=he*1
								 n=n*1															 
						if he=n then 
	
		%>
                          <option value="<%response.Write(n)%>" selected>
                            <%response.Write(h_exibe)%>
                            </option>
                          <% else%>
                          <option value="<%response.Write(n)%>">
                            <%response.Write(h_exibe)%>
                            </option>
                          <% end if 
							next
							%>
                        </select>
                        :
                        <select name="min_de" id="min_de" class="select_style" <%response.Write(disable)%>>
                          <% 
							For p =0 to 59
							 	mn=mn*1
								 p=p*1	
								if p<10 then
								 m_exibe="0"&p
								else
								 m_exibe=p						 
								end if	
							 	mn=mn*1
								 p=p*1															 						 
							 if mn=p then %>
                          <option value="<%response.Write(p)%>" selected>
                            <%response.Write(m_exibe)%>
                            </option>
                          <% else

										%>
                          <option value="<%response.Write(p)%>">
                            <%response.Write(m_exibe)%>
                            </option>
                          <% end if 
							next
							%>
                        </select>
                        <%end if%></td>
                      <td height="10" align="center">
					  <% if ori = "C" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Tipo_Entrevista where TP_Entrevista = "&tp_entrevista
		RS.Open SQL, CON0
					
		response.Write(RS("TX_Descricao"))  %>
                        <input name="tipo" type="hidden" value="<%response.Write(tp_entrevista)%>">
                        <% 
		else %>
                        <select name="tipo" id="tipo" class="select_style" >
                          <% 
					Set RS = Server.CreateObject("ADODB.Recordset")
					SQL = "SELECT * FROM TB_Tipo_Entrevista order by NU_Prioridade_Combo"
					RS.Open SQL, CON0
				while not RS.EOF	
				
					tipo_entrevista = RS("TP_Entrevista")
					tx_descricao = RS("TX_Descricao")			
					
	
			tipo_entrevista=tipo_entrevista*1
			tp_entrevista=tp_entrevista*1
			if tipo_entrevista = tp_entrevista then
						  %>
                          <option value="<%response.Write(tipo_entrevista)%>" selected>
                            <%response.Write(tx_descricao)%>
                            </option>
                          <%			
			else							
						  %>
                          <option value="<%response.Write(tipo_entrevista)%>">
                            <%response.Write(tx_descricao)%>
                            </option>
                          <% end if
        RS.MOVENEXT
        Wend
        %>
                        </select>
                        <%end if%></td>
                       <%if ori="I" then
					   		colspan = "3"
						 else
					   		colspan = "2"						 
						 end if	 %>
                      <td height="10" colspan="<%response.Write(colspan)%>" align="center"><% if ori = "C" then

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Usuario Where CO_Usuario = "&ag_entrevista
		RS1.Open SQL1, CON
					
		response.Write(RS1("NO_Usuario"))  %>
                        <input name="agendado" type="hidden" value="<%response.Write(ag_entrevista)%>">
                        <% 
else %>
                        <select name="agendado" id="agendado" class="select_style">
                          <% 
				Set RS1 = Server.CreateObject("ADODB.Recordset")
				SQL1 = "SELECT * FROM TB_Usuario Where CO_Usuario <> 0 order by NO_Usuario"
				RS1.Open SQL1, CON
				while not RS1.EOF	
				
					co_atendido = RS1("CO_Usuario")
					nome_atendido = RS1("NO_Usuario")					
					ag_entrevista=ag_entrevista*1
					co_atendido=co_atendido*1
			if co_atendido = ag_entrevista then
						  %>
                          <option value="<%response.Write(co_atendido)%>" selected>
                            <%response.Write(nome_atendido)%>
                            </option>
                          <%			
			else							
						  %>
                          <option value="<%response.Write(co_atendido)%>">
                            <%response.Write(nome_atendido)%>
                            </option>
                          <% end if
        RS1.MOVENEXT
        Wend
        %>
                        </select>
                        <%end if%></td>
                       <%if ori="A" then%>                        
                      <td align="center"><select name="status" id="status" class="select_style">
                         					
						<%
                                st_entrevista=st_entrevista*1
                        if st_entrevista = 3 then
                                      %>
                                      <option value="1">
                                       Atendida
                                        </option>                                          
                                      <option value="3" selected>
                                        Pendente
                                        </option>
                                      <option value="2">
                                       Cancelado
                                        </option>                            
                                      <%			
                        elseif st_entrevista = 2 then							
                                      %>
                                      <option value="1">
                                       Atendida
                                        </option>                                       
                                      <option value="3">
                                        Pendente
                                        </option>                          
                                      <option value="2" selected>
                                       Cancelado
                                        </option>
                       <% elseif st_entrevista = 1 then							
                                      %>
                                      <option value="1" selected>
                                       Atendida
                                        </option>                                       
                                      <option value="3">
                                        Pendente
                                        </option>                          
                                      <option value="2">
                                       Cancelado
                                        </option>                                        
                      <% end if
                    %>
                        </select></td>
                       <%elseif ori="C" then%>
                      <td align="center">
                         					
					<%
					st_entrevista=st_entrevista*1
					if st_entrevista = 3 then
								  %>
									Pendente
								  <%			
					elseif st_entrevista = 2 then							
								  %>
								   Cancelado
                     <%else%>
                     Atendida             
                          <% end if%></td>
                       <%end if%>                        
					    
                    </tr>
                  </table></td>
                </tr>
                
<% if ori="C" then  

	'if tp_entrevista>1 then

		Set RSC = Server.CreateObject("ADODB.Recordset")
		SQLC = "SELECT * FROM TB_Entrevistas_Conteudo WHERE CO_Matricula ="& cod&" AND (DA_Entrevista=#"&da_entrevista_cons&"# AND mid(HO_Entrevista,1,16)=#12/30/1899 "&ho_entrevista_cons&"#)" 
		RSC.Open SQLC, CON4
		
		if RSC.EOF then
			conteudo = ""
		else 
			conteudo = RSC("TX_Conteudo")
		end if

%>
                <tr class="tb_subtit">
                  <td height="10" colspan="8">Conte&uacute;do</td>
                </tr>
                <tr >
                  <td height="10" colspan="8"><textarea name="conteudo" cols="190" rows="20" class="textInput" id="conteudo" <%response.Write(disable)%>><%response.Write(conteudo)%></textarea></td>
                </tr>

<%
	'end if

else%>             
                <tr class="tb_subtit"> 
                  <td height="10" colspan="8">Participantes 
                    <input name="data_altera" type="hidden" id="data_altera" value="<%=data_altera%>">
                    </td>
                </tr>
                <tr > 
                  <td height="10" colspan="8"><input name="participantes" type="text" class="textInput" id="participantes" value="<%response.Write(partic_entrevista)%>" size="195" maxlength="255" ></td>
                </tr>
<%end if%>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2"><div align="right"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></div></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td colspan="2"><div align="center"> 
                <table width="1000" border="0" align="center" cellspacing="0">
                  <tr> 
                    <td height="24" colspan="3">
<hr></td>
                  </tr>
                  <tr> 
                    <td width="33%"> 
                      <div align="center"> 
                        <input name="alterar" type="submit" class="botao_cancelar" id="alterar" onClick="MM_goToURL('parent','resumo.asp?or=2');return document.MM_returnValue" value="Voltar">

                      </div></td>
                    <td width="34%">&nbsp;</td>
                    <td width="33%"> <div align="center">
                    <%	
					'if ori = "C" and tp_entrevista=1 then 
					
					'else%>
                        <input name="Submit" type="submit" class="botao_prosseguir" value="Confirmar">
					<% 'end if%>	
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