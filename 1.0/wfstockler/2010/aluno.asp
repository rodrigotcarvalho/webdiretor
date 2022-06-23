<!--#include file="inc/caminhos.asp"-->
<!--#include file="inc/funcoes.asp"-->
<!--#include file="inc/funcoes2.asp"-->


<%
nivel=0
tp=session("tp")
nome = session("nome") 
co_user = session("co_user")

 	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0 
	
	Set CON = Server.CreateObject("ADODB.Connection")
 	ABRIR = "DBQ="& CAMINHO_wf& ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1
	
				Set CON7 = Server.CreateObject("ADODB.Connection") 
		ABRIR7 = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON7.Open ABRIR7

opt=request.QueryString("opt")

If tp= "A" then
co_aluno=co_user
msg="Seja bem-vindo"
else
	if opt="d" then
		SQL1 = "select * from TB_RespxAluno where CO_Usuario = " & co_user 
		set RS1 = CON.Execute (SQL1)

		co_aluno= RS1("CO_Aluno")

		msg="Aluno Vinculado"

	else
		co_aluno= request.form("co_aluno")
		msg="Aluno Selecionado"
	end if
end if
	SQL2 = "select * from TB_Usuario where CO_Usuario = " & co_aluno 
	set RS2 = CON.Execute (SQL2)
	
nome_aluno= RS2("NO_Usuario")

%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Web Família</title>
<link href="estilo.css" rel="stylesheet" type="text/css" />
<script type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
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

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
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
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
//-->
</script>
</head>

<body onLoad="MM_preloadImages(<%response.Write(swapload)%>)">
<%if tp="R" and opt="sa" then%>
<%end if%>
<table width="1000" height="1078" border="0" align="center" cellpadding="0" cellspacing="0" class="tabela_principal">
  <%
			ano = DatePart("yyyy", now)
			mes = DatePart("m", now) 
			dia = DatePart("d", now) 
			hora = DatePart("h", now) 
			min = DatePart("n", now) 
select case mes
 case 1 
 mes = "janeiro"
 case 2 
 mes = "fevereiro"
 case 3 
 mes = "março"
 case 4
 mes = "abril"
 case 5
 mes = "maio"
 case 6 
 mes = "junho"
 case 7
 mes = "julho"
 case 8 
 mes = "agosto"
 case 9 
 mes = "setembro"
 case 10 
 mes = "outubro"
 case 11 
 mes = "novembro"
 case 12 
 mes = "dezembro"
end select

data = dia &" / "& mes &" / "& ano
data= FormatDateTime(data,1) 			

			horario = hora & ":"& min%>
  <tr>
    <td height="998"><table width="200" height="998" border="0" cellpadding="0" cellspacing="0">
        <!--DWLayoutTable-->
        <tr valign="bottom"> 
          <td height="90" colspan="3"> 
            <%call cabecalho(nivel)%>
          </td>
        </tr>
        <tr class="tabela_menu"> 
          <td width="172" rowspan="4" valign="top" class="tabela_menu"> <p>&nbsp;</p>
            <% call menu_lateral(nivel)%>
            <p>&nbsp;</p></td>
          <td width="640" height="12" valign="bottom" nowrap="nowrap">
<p class="style1">&nbsp;&nbsp;Ol&aacute; <span class="style2"> 
              <%response.Write(nome)%>
              </span> , &uacute;ltimo acesso dia 
              <% Response.Write(session("dia_t")) %>
              &agrave;s 
              <% Response.Write(session("hora_t")) %>
            </p></td>
          <td width="188"><p align="right" class="style1"> 
              <%response.Write(data)%>
            </p></td>
        </tr>
        <tr class="tabela_menu"> 
          <td height="5" colspan="2"><p><img src="img/linha-pontilhada_grande.gif" alt="" width="828" height="5" /></p></td>
        </tr>
      <tr class="tabela_menu">
        <td height="19" colspan="2">&nbsp;</td>
      </tr>		
        <tr class="tabela_menu"> 
          <td height="832" colspan="2" valign="top"> <p align="left"><img src="img/inicial.jpg" width="700" height="30"></p>
            <table width="800" border="0" cellspacing="0" cellpadding="0">
              <tr> 
                <td width="10" height="30">&nbsp;</td>
                <td height="30" colspan="7" valign="top"><font class="style3">O 
                  Web Fam&iacute;lia disponibilizar&aacute; acesso &agrave;s informa&ccedil;&otilde;es 
                  do aluno que estiver selecionado abaixo:</font></td>
              </tr>
              <tr> 
                <td width="10">&nbsp;</td>
                <td width="20">&nbsp;</td>
                <td width="80"><div align="center"><font class="style3">MATR&Iacute;CULA</font></div></td>
                <td width="245"><font class="style3">NOME</font></td>
                <td width="80"> <div align="center"><font class="style3">UNIDADE</font></div></td>
                <td width="195"><font class="style3">CURSO</font></td>
                <td width="60"> <div align="center"><font class="style3">TURMA</font></div></td>
                <td width="110"> <div align="center"><font class="style3">NASCIMENTO</font></div></td>
              </tr>
              <%
 	

	SQL2 = "select * from TB_Alunos where CO_Matricula = " & co_aluno 
	set RS2 = CON1.Execute (SQL2)
	
nome_aluno= RS2("NO_Aluno")
		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Contatos where CO_Matricula = " & co_aluno &" AND TP_Contato='ALUNO'"
		RS7.Open SQL7, CON7

nascimento = RS7("DA_Nascimento_Contato")

dados_dtd= split(nascimento, "/" )
dia_de= dados_dtd(0)
mes_de= dados_dtd(1)
ano_de= dados_dtd(2)

if dia_de<10 then
dia_de="0"&dia_de
end if
if mes_de<10 then
mes_de="0"&mes_de
end if

nascimento=dia_de&"/"&mes_de&"/"&ano_de

	SQL3 = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND CO_Matricula ="& co_aluno
	set RS3 = CON1.Execute (SQL3)

nu_unidade= RS3("NU_Unidade")
co_curso= RS3("CO_Curso")
co_etapa= RS3("CO_Etapa")
co_turma= RS3("CO_Turma")

call GeraNomes("PORT",nu_unidade,co_curso,co_etapa,CON0)
no_unidade = session("no_unidade")
no_curso = session("no_curso")
no_etapa = session("no_etapa")
prep_curso=session("prep_curso")
local= no_etapa&" "&prep_curso&" "&no_curso
%>
              <tr> 
                <td>&nbsp;</td>
                <td> <input name="radiobutton" type="radio" value="radiobutton" checked> 
                </td>
                <td><div align="center"><font class="style1"> 
                    <%response.write(co_aluno)%>
                    </font> </div></td>
                <td> <font class="style1"> 
                  <%response.write(nome_aluno)%>
                  </font> </td>
                <td><div align="center"><font class="style1"> 
                    <%response.write(no_unidade)%>
                    </font></div></td>
                <td><font class="style1"> 
                  <%response.write(local)%>
                  </font></td>
                <td><div align="center"><font class="style1"> 
                    <%response.write(co_turma)%>
                    </font></div></td>
                <td><div align="center"><font class="style1"> 
                    <%response.write(nascimento)%>
                    </font></div></td>
              </tr>
            </table></td>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td width="1000" height="40" valign="top"><img src="img/rodape.jpg" width="1000" height="40" /></td>
  </tr>
</table>
</body>
</html>
