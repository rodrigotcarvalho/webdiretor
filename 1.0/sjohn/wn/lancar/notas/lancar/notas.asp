<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes3.asp"-->

<%
opt=request.QueryString("opt")
obr = request.QueryString("obr")
autoriza=session("autoriza")
nvg = session("chave")
ano_letivo = request.QueryString("ano")
co_usr = session("co_user")
grupo=session("grupo")
chave=nvg
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
nivel=4

nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

if opt = "err6" then

hp= request.QueryString("hp")

dados= split(obr, "?" )
co_materia = dados(0)
unidades= dados(1)
grau= dados(2)
serie= dados(3)
turma= dados(4)
periodo = dados(5)
ano_letivo = dados(6)
co_prof = dados(7)
co_usr = session("co_usr")


alt= split(hp, "_" )
calc=alt(1)
qerrou= alt(2)
errou = alt(3)

if calc=0 then

else
		Set CONer= Server.CreateObject("ADODB.Connection") 
		ABRIRer = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONer.Open ABRIRer
		
		Set RSer  = Server.CreateObject("ADODB.Recordset")
		SQL_er  = "Select * from TB_Matriculas WHERE CO_Matricula = "& calc &" AND NU_Ano = "&ano_letivo
		Set RSer  = CONer.Execute(SQL_er )
		


errante = RSer("NU_Chamada")
end if

select case errou
case "f"
errado ="Faltas"
case "pt"
errado ="Peso dos testes (vazio)"
case "pp"
errado ="Peso das Provas (vazio)"
case "t1"
errado ="T1"
case "t2"
errado ="T2"
case "t3"
errado ="T3"
case "t4"
errado ="T4"
case "mt"
errado ="T1, T2, T3 ou T4 (a soma dos testes deve ser menor que 100)"
case "p1"
errado ="P1"
case "p1b"
errado ="P1 (Deve ser menor ou igual a 80)"
case "p2"
errado ="P2"
case "simul"
errado ="S (Deve ser menor ou igual a 20)"
case "p3"
errado ="P3"
case "av1"
errado ="AV1"
errou = "va_av1"
case "av2"
errado ="AV2"
errou = "va_av2"
case "av3"
errado ="AV3"
errou = "va_av3"
case "av4"
errado ="AV4"
errou = "va_av4"
case "av5"
errado ="AV5"
errou = "va_av5"
case "sim"
errado ="SIM"
errou = "va_sim"
case "simv"
errado ="SIM (Deve ser menor ou igual a 10)"
errou = "va_sim"
case "bat"
errado ="BAT"
errou = "va_bat"
case "batv"
errado ="BAT (Deve ser menor ou igual a 5)"
errou = "va_bat"

case "bon"
errado ="Bon"
case "rec"
errado ="Rec"
case "recp"
errado ="Rec s&oacute; pode ser lan&ccedil;ada nos per&iacute;odos 1 e 2"
errou = "va_rec"
case "rec70"
errado ="Rec s&oacute; pode ser lan&ccedil;ada se M2<70"
errou = "va_rec"
end select

elseif opt="ok" or  opt= "vt" or opt="cln" then
dados= split(obr, "?" )
co_materia = dados(0)
unidades= dados(1)
grau= dados(2)
serie= dados(3)
turma= dados(4)
periodo = dados(5)
ano_letivo = dados(6)
co_prof = dados(7)
co_usr = session("co_usr")

elseif opt="cgp" then
co_materia = request.QueryString("d")
unidades= request.QueryString("u")
grau= request.QueryString("c")
serie= request.QueryString("e")
turma= request.QueryString("t")
periodo = request.QueryString("p")
co_prof = request.querystring("pr")
co_usr = session("co_usr")

else
unidades= request.form("unidade")
grau= request.form("curso")
serie= request.form("etapa")
turma= request.form("turma")
co_materia = request.form("mat_prin")
periodo = request.form("periodo")
co_prof = Session("co_prof")
co_usr = session("co_usr")
grava_nota= unidades&"?"&grau&"?"&serie&"?"&turma
session("grava_nota")=grava_nota
end if
session("co_materia")=co_materia
session("unidades")=unidades
session("grau")=grau
session("serie")=serie
session("turma")=turma
session("periodo")=periodo



if serie = "999999" then
		response.redirect("tabelas2.asp?opt=err1&or=1&dd="&ano_letivo&"_"&grau&"_"&unidades&"_"&co_prof)
ELSEif co_materia = "999999" then
		response.redirect("tabelas2.asp?opt=err2&or=1&dd="&ano_letivo&"_"&grau&"_"&unidades&"_"&co_prof)
ELSEif turma = "999999" then
	response.redirect("tabelas2.asp?opt=err3&or=1&dd="&ano_letivo&"_"&grau&"_"&unidades&"_"&co_prof)

else

obr=co_materia&"?"&unidades&"?"&grau&"?"&serie&"?"&turma&"?"&periodo&"?"&ano_letivo&"?"&co_prof

				Set CONg = Server.CreateObject("ADODB.Connection") 
		ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONg.Open ABRIRg

				Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidades &" AND CO_Curso = '"& grau &"' AND CO_Etapa = '"& serie &"' AND CO_Turma = '"& turma &"'"
		Set RS = CONg.Execute(CONEXAO)


if RS.EOF then
response.Write("<div align=center><font size=2 face=Courier New, Courier, mono  color=#990000><b>Esta turma n�o est� dispon�vel no momento</b></font><br")
response.Write("<font size=2 face=Courier New, Courier, mono  color=#990000><a href=javascript:window.history.go(-1)>voltar</a></font></div>")

else
nota = RS("TP_Nota")
coordenador = RS("CO_Cord")
end if

if co_materia = "RED2" AND (nota = "TB_NOTA_F" or nota = "TB_NOTA_K" or nota = "TB_NOTA_M") THEN
	nota = "TB_NOTA_V"
END IF

 call navegacao (CON,chave,nivel)
navega=Session("caminho")
trava=session("trava")
%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../../../../estilos.css" type="text/css">
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

function mudar_cor_focus(celula){
   celula.style.backgroundColor="#D8FF9D"

}
function mudar_cor_blur_par(celula){
   celula.style.backgroundColor="#FFFFFF"
} 
function mudar_cor_blur_impar(celula){
   celula.style.backgroundColor="#FFFFE1"
} 
function mudar_cor_blur_erro(celula){
   celula.style.backgroundColor="#CC0000"
}  
function checksubmit()
{
<%
if nota<>"TB_NOTA_V" then
%>
// if (document.nota.pt.value == "")
//  {    alert("Por favor digite um peso para os Testes!")
//    document.nota.pt.focus()
//    return false
//  }
  if (isNaN(document.nota.pt.value))
  {    alert("O peso dos Testes deve ser um n�mero!")
    document.nota.pt.focus()
    return false
  }  

<%
end if
if nota="TB_NOTA_F" then%>
//    if (document.nota.pp1.value == "")
//  {    alert("Por favor digite um peso para a Prova1!")
//    document.nota.pp1.focus()
//    return false
//  }
  if (isNaN(document.nota.pp1.value))
  {    alert("O peso da Prova 1 deve ser um n�mero!")
    document.nota.pp1.focus()
    return false
  }
//    if (document.nota.pp2.value == "")
//  {    alert("Por favor digite um peso para a Prova 2!")
//    document.nota.pp2.focus()
//    return false
//  }  
  if (isNaN(document.nota.pp2.value))
  {    alert("O peso das Prova 2 deve ser um n�mero!")
    document.nota.pp2.focus()
    return false
  }  
<%

elseif nota<>"TB_NOTA_V" then%>
//    if (document.nota.pp.value == "")
//  {    alert("Por favor digite um peso para as Provas!")
//    document.nota.pp.focus()
//    return false
//  }
  if (isNaN(document.nota.pp.value))
  {    alert("O peso das Provas deve ser um n�mero!")
    document.nota.pp.focus()
    return false
  }
<%end if%> 

  return true
}
//-->
</script>
<script language="javascript"> 
  
    function keyPressed(TB, e, max_right, max_bottom)  
    { 
        if (e.keyCode == 40 || e.keyCode == 13) { // arrow down 
            if (TB.split("c")[0] < max_bottom) 
            document.getElementById(eval(TB.split("c")[0] + '+1') + 'c' + TB.split("c")[1]).focus(); 
            if (TB.split("c")[0] == max_bottom) 
            document.getElementById(1 + 'c' + TB.split("c")[1]).focus();


        } 
  
        if (e.keyCode == 38) { // arrow up 
            if(TB.split("c")[0] > 1) 
            document.getElementById(eval(TB.split("c")[0] + '-1') + 'c' + TB.split("c")[1]).focus(); 
            if (TB.split("c")[0] == 1) 
            document.getElementById(max_bottom + 'c' + TB.split("c")[1]).focus(); 
		
        } 
  
        if (e.keyCode == 37) { // arrow left 
            if(TB.split("c")[1] > 1) 
            document.getElementById(TB.split("c")[0] + 'c' + eval(TB.split("c")[1] + '-1')).focus();             
            if (TB.split("c")[1] == 1) 
            document.getElementById(TB.split("c")[0] + 'c' + max_right).focus(); 

		}   
  
        if (e.keyCode == 39) { // arrow right 
            if(TB.split("c")[1] < max_right) 
            document.getElementById(TB.split("c")[0] + 'c' + eval(TB.split("c")[1] + '+1')).focus();  
            if (TB.split("c")[1] == max_right) 
            document.getElementById(TB.split("c")[0] + 'c' + 1).focus(); 

		}                  
    } 
  
</script> 
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" background="../../../../img/fundo.gif" marginheight="0">
<%IF imprime="1"then
else
 call cabecalho (nivel) 
 end if%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
    <td height="10" valign="top" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
    <%

			Set Conecta= Server.CreateObject("ADODB.Connection") 
		ABRIRgn = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		Conecta.Open ABRIRgn

call GeraNomes(co_materia,unidades,grau,serie,Conecta)

no_materia= session("no_materia")
no_unidades= session("no_unidades")
no_grau= session("no_grau")
no_serie= session("no_serie")


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


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE CO_Professor= "& co_prof &"AND NU_Unidade = "& unidades &" AND CO_Curso = '"& grau &"' AND CO_Etapa = '"& serie &"' AND CO_Turma = '"& turma &"' AND CO_Materia_Principal = '"& co_materia &"'"
		Set RS = CON.Execute(CONEXAO)
ST_Per_1 = RS("ST_Per_1")
ST_Per_2 = RS("ST_Per_2")
ST_Per_3 = RS("ST_Per_3")
ST_Per_4 = RS("ST_Per_4")
ST_Per_5 = RS("ST_Per_5")
ST_Per_6 = RS("ST_Per_6")

tp = session("tp")
%>
            <%if opt = "ok" then%>
            <tr>         
    <td height="10" valign="top"> 
      <%
		call mensagens(nivel,622,2,0)
%>
      <div align="center"></div></td>
            </tr>			
            <%elseif opt= "err6" then %>
            <tr> 
    <td height="10" valign="top"> 
      <%
	call mensagens(nivel,620,1,errante)
%>
</td>
            </tr>
            <%end if
%>
            <% IF trava="s" or (co_usr<>coordenador AND grupo="COO") then%>
            <tr>     
    <td height="10" valign="top"> 
      <%	call mensagens(nivel,9701,0,0) %>
</td>
            </tr>
            <%elseif (periodo = 1 and ST_Per_1="x") OR (periodo = 2 and ST_Per_2="x") OR (periodo = 3 and ST_Per_3="x") OR (periodo = 4 and ST_Per_4="x") OR (periodo = 5 and ST_Per_5="x") OR (periodo = 6 and ST_Per_6="x") then%>
            <tr> 
    <td height="10" valign="top"> 
      <%
		call mensagens(nivel,624,1,0)
%>
</td>
            </tr>

            <% end if%>
<%if opt= "cln" then %>
            <tr> 
    <td height="10" valign="top"> 
      <%
	call mensagens(nivel,621,0,0)
%>
</td>
            </tr>
            <% end if%>						
	            <tr> 
    <td height="10" valign="top"> 
      <%	call mensagens(nivel,645,0,0) 

%>
    </td>
            </tr>			
            <tr class="tb_tit"> 
              
    <td height="15" class="tb_tit">&nbsp;Grade de Aulas</td>
            </tr>
            <tr> 
    <td height="36" valign="top"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td width="166" class="tb_subtit"> 
            <div align="center"><strong>UNIDADE 
              </strong></div></td>
          <td width="166" class="tb_subtit"> 
            <div align="center"><strong>CURSO 
              </strong></div></td>
          <td width="166" class="tb_subtit"> 
            <div align="center"><strong>ETAPA 
              </strong></div></td>
          <td width="166" class="tb_subtit"> 
            <div align="center"><strong>TURMA 
              </strong></div></td>
          <td width="170" class="tb_subtit"> 
            <div align="center"><strong>DISCIPLINA</strong></div></td>
          <td width="166" class="tb_subtit"> 
            <div align="center"><strong>PER&Iacute;ODO 
              </strong></div></td>
        </tr>
        <tr> 
          <td width="166"> 
            <div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%response.Write(no_unidades)%>
              </font></div></td>
          <td width="166"> 
            <div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%response.Write(no_grau)%>
              </font></div></td>
          <td width="166"> 
            <div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%
response.Write(no_serie)%>
              </font></div></td>
          <td width="166"> 
            <div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%
response.Write(turma)%>
              </font></div></td>
          <td width="170"> 
            <div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%

response.Write(no_materia)%>
              </font> </div></td>
          <td width="166"> 
            <div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%
		Set RSper = Server.CreateObject("ADODB.Recordset")
		SQLper = "SELECT * FROM TB_Periodo where NU_Periodo= "&periodo
		RSper.Open SQLper, Conecta

NO_Periodo= RSper("NO_Periodo")
response.Write(NO_Periodo)%>
              </font> </div></td>
        </tr>
      </table></td>
            </tr>
      <tr> 
        
    <td valign="top"> 
      <%

co_usr=co_usr*1
coordenador=coordenador*1
autoriza=autoriza*1
if opt="cln" then
	if nota ="TB_NOTA_A" then
	CAMINHOn = CAMINHO_na
	Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"A","cln",0)
	else
		if nota="TB_NOTA_B" then
		CAMINHOn = CAMINHO_nb
		Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"B","cln",0)
		else
			if nota ="TB_NOTA_C" then
			CAMINHOn = CAMINHO_nc
			Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"C","cln",0)
			else
				if nota ="TB_NOTA_D" then
				CAMINHOn = CAMINHO_nd
				Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"D","cln",0)
				else
					if nota ="TB_NOTA_E" then
					CAMINHOn = CAMINHO_ne
					Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"E","cln",0)
					else
						if nota ="TB_NOTA_F" then
						CAMINHOn = CAMINHO_nf
						Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"F","cln",0)
						else
							if nota ="TB_NOTA_V" then
							CAMINHOn = CAMINHO_nv
							Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"V","cln",0)
							else
								if nota ="TB_NOTA_K" then
								CAMINHOn = CAMINHO_nk
								Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"K","cln",0)
								else
									if nota ="TB_NOTA_L" then
									CAMINHOn = CAMINHO_nl
									Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"L","cln",0)
									else
										if nota ="TB_NOTA_M" then
										CAMINHOn = CAMINHO_nm
										Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"M","cln",0)
										else
										response.Write("ERRO")
										End if
									End if
								End if
							End if
						End if
					End if
				End if
			End if
		end if
	end if

elseif ((periodo = 1 and ST_Per_1="x") OR (periodo = 2 and ST_Per_2="x") OR (periodo = 3 and ST_Per_3="x") AND co_usr<>coordenador) OR trava="s" then
	if nota ="TB_NOTA_A" then
	CAMINHOn = CAMINHO_na
	Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"A","blq",0)
	else
		if nota="TB_NOTA_B" then
		CAMINHOn = CAMINHO_nb
		Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"B","blq",0)
		else
			if nota ="TB_NOTA_C" then
			CAMINHOn = CAMINHO_nc
			Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"C","blq",0)
			else
				if nota ="TB_NOTA_D" then
				CAMINHOn = CAMINHO_nd
				Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"D","blq",0)
				else
					if nota ="TB_NOTA_E" then
					CAMINHOn = CAMINHO_ne
					Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"E","blq",0)
					else
						if nota ="TB_NOTA_F" then
						CAMINHOn = CAMINHO_nf
						Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"F","blq",0)
						else
							if nota ="TB_NOTA_V" then
							CAMINHOn = CAMINHO_nv
							Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"V","blq",0)
							else
								if nota ="TB_NOTA_K" then
								CAMINHOn = CAMINHO_nk
								Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"K","blq",0)
								else
									if nota ="TB_NOTA_L" then
									CAMINHOn = CAMINHO_nl
									Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"L","blq",0)
									else
										if nota ="TB_NOTA_M" then
										CAMINHOn = CAMINHO_nm
										Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"M","blq",0)
										else
										response.Write("ERRO")
										End if
									End if
								End if
							End if
						End if
					End if
				End if
			End if
		end if
	end if
ELSEIF (co_usr=coordenador or autoriza=5) AND trava<>"s" then
	if nota ="TB_NOTA_A" then
	CAMINHOn = CAMINHO_na
	Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"A","edit",0)
	else
		if nota="TB_NOTA_B" then
		CAMINHOn = CAMINHO_nb
		Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"B","edit",0)
		else
			if nota ="TB_NOTA_C" then
			CAMINHOn = CAMINHO_nc
			Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"C","edit",0)
			else
				if nota ="TB_NOTA_D" then
				CAMINHOn = CAMINHO_nd
				Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"D","edit",0)
				else
					if nota ="TB_NOTA_E" then
					CAMINHOn = CAMINHO_ne
					Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"E","edit",0)
					else
						if nota ="TB_NOTA_F" then
						CAMINHOn = CAMINHO_nf
						Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"F","edit",0)
						else
							if nota ="TB_NOTA_V" then
							CAMINHOn = CAMINHO_nv
							Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"V","edit",0)
							else
								if nota ="TB_NOTA_K" then
								CAMINHOn = CAMINHO_nk
								Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"K","edit",0)
								else
									if nota ="TB_NOTA_L" then
									CAMINHOn = CAMINHO_nl
									Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"L","edit",0)
									else
										if nota ="TB_NOTA_M" then
										CAMINHOn = CAMINHO_nm
										Call notas(CAMINHO_al,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,"M","edit",0)
										else
										response.Write("ERRO")
										End if
									End if
								End if
							End if
						End if
					End if
				End if
			End if
		end if
	end if	
end if	
 %>
    </td>
      </tr>
      <%	
    Set RS = Nothing
	    Set RS2 = Nothing
		    Set RS3 = Nothing
End if
%>
      <tr>      
    <td height="40" valign="top"> <img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
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