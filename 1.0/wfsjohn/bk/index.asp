<!--#include file="inc/caminhos.asp"-->
<!--#include file="inc/funcoes.asp"-->
<!--#include file="inc/funcoes2.asp"-->


<%
nivel=2
tipo_arquivo=2
session("tipo_arquivo") = tipo_arquivo
'tp=session("tp")
'nome = session("nome") 
'co_user = session("co_user")
'opt=request.QueryString("opt")

'if opt="1" then
'periodo_check=request.form("periodo")
'cod= Session("aluno_selecionado")
'else
'cod= Session("aluno_selecionado")
'periodo_check=1
'end if
'cod= Session("aluno_selecionado")

'obr=cod&"?"&periodo_check

 	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0

 	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1
	
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2	



	SQL2 = "select * from TB_Usuario where CO_Usuario = 116"
	set RS2 = CON.Execute (SQL2)
	
nome_aluno= RS2("NO_Usuario")

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Matriculas WHERE NU_Ano=2008 AND CO_Matricula =116"
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



%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Web Fam�lia</title>
<link href="estilo.css" rel="stylesheet" type="text/css" />
<script type="text/JavaScript">
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

function submitfuncao()  
{
   var f=document.forms[0]; 
      f.submit(); 
}
//-->
</script>
</head>

<body onLoad="MM_preloadImages(<%response.Write(swapload)%>)">
<form action="docs/avcirc/index.asp?opt=1" method="post"><table width="1000" height="1039" border="0" align="center" cellpadding="0" cellspacing="0" class="tabela_principal">
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
 mes = "mar�o"
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
          <td height="120" colspan="3"> 
              <%call cabecalho(nivel)%>
            </td>
          </tr>
          <tr class="tabela_menu"> 
            <td width="172" height="144" rowspan="4" valign="top" class="tabela_menu"><p><img src="img/baner_resto.jpg" width="171" height="19" /></p>
              <% call menu_lateral(nivel)%>
              <p>&nbsp;</p></td>
            <td width="640" height="12" nowrap="nowrap"><p class="style1">&nbsp;&nbsp;Ol&aacute; 
                <span class="style2">
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
            <td height="832" colspan="2" valign="top"> <div align="left"><img src="img/avcirc.jpg" width="700" height="30"> 
              </div>
              <table width="100%" border="0" align="center" cellspacing="0" class="tb_corpo">
                <tr> 
                  <td colspan="2" > <table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <%
'		Set RS_tb = Server.CreateObject("ADODB.Recordset")
'		SQL_tb = "SELECT * FROM TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso&"' AND CO_Etapa ='"& etapa&"' AND CO_Turma ='"& turma&"'"
'		RS_tb.Open SQL_tb, CON2

'notaFIL=RS_tb("TP_Nota")



'if notaFIL ="TB_NOTA_A" then
'CAMINHOn = CAMINHO_na

'elseif notaFIL="TB_NOTA_B" then
'	CAMINHOn = CAMINHO_nb

'elseif notaFIL ="TB_NOTA_C" then
'		CAMINHOn = CAMINHO_nc
'else
'		response.Write("ERRO")
'end if
		'response.Write "SELECT * FROM TB_Documentos where TP_Doc= "&tipo_arquivo&" AND (((Unidade="&unidade&") AND (Curso="&curso&") AND  (Etapa="&etapa&") AND (Turma="&turma&")) OR ((Unidade="&unidade&") AND (Curso Is Null) AND  (Etapa Is Null) AND (Turma Is Null)) OR ((Unidade="&unidade&") AND (Curso="&curso&") AND  (Etapa Is Null) AND (Turma Is Null)) OR ((Unidade="&unidade&") AND (Curso="&curso&") AND  (Etapa="&etapa&") AND (Turma Is Null)) OR  ((Unidade Is Null) AND (Curso Is Null) AND  (Etapa Is Null) AND (Turma Is Null))) order by DA_Doc Desc"

'		SQL_doc = "SELECT * FROM TB_Documentos where TP_Doc= "&tipo_arquivo&" AND (((Unidade="&unidade&") AND (Curso="&curso&") AND  (Etapa="&etapa&") AND (Turma="&turma&")) OR ((Unidade="&unidade&") AND (Curso Is Null) AND  (Etapa Is Null) AND (Turma Is Null)) OR ((Unidade="&unidade&") AND (Curso="&curso&") AND  (Etapa Is Null) AND (Turma Is Null)) OR ((Unidade="&unidade&") AND (Curso="&curso&") AND  (Etapa="&etapa&") AND (Turma Is Null)) OR  ((Unidade Is Null) AND (Curso Is Null) AND  (Etapa Is Null) AND (Turma Is Null))) order by DA_Doc Desc"
	
		Set RS_doc = Server.CreateObject("ADODB.Recordset")
'		SQL_doc = "SELECT * FROM TB_Documentos where TP_Doc= "&tipo_arquivo&" AND ((Unidade="&unidade&") AND (Curso="&curso&") AND  (Etapa="&etapa&") AND (Turma Is Null)) order by DA_Doc Desc"
		SQL_doc = "SELECT * FROM TB_Documentos where TP_Doc= "&tipo_arquivo&" AND (((Unidade='"&unidade&"') AND (Curso='"&curso&"') AND  (Etapa='"&etapa&"') AND (Turma='"&turma&"')) OR ((Unidade='"&unidade&"') AND (Curso Is Null) AND  (Etapa Is Null) AND (Turma Is Null)) OR ((Unidade='"&unidade&"') AND (Curso='"&curso&"') AND  (Etapa Is Null) AND (Turma Is Null)) OR ((Unidade='"&unidade&"') AND (Curso='"&curso&"') AND  (Etapa='"&etapa&"') AND (Turma Is Null)) OR  ((Unidade Is Null) AND (Curso Is Null) AND  (Etapa Is Null) AND (Turma Is Null))) order by DA_Doc Desc"
		RS_doc.Open SQL_doc, CON

if RS_doc.eof then
%>
                      <tr class="<%response.write(cor)%>"> 
                        <td colspan="2" valign="top"> <div align="center"><font class="style1"> 
                          <%response.Write("<br><br><br><br><br>N�o existem documentos para este segmento!")%>
                          </font></div></td>
                      </tr>

<%else
check=2
ordem=0
while not RS_doc.eof

 if check mod 2 =0 then
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if

%>

                      <tr class="<%response.write(cor)%>"> 
                        <%
ordem=ordem+1						
tit1=RS_doc("TI1_Doc")
nome_arq=RS_doc("NO_Doc")
extensao_arq = Array()
extensao_arq= split(nome_arq, "." )
extensao= extensao_arq(ubound(extensao_arq))
nome_sessao="arq_"&ordem
session(nome_sessao)=nome_arq

select case extensao

case "doc"
icone="word"

case "xls"
icone="excel"

case "pdf"
icone="pdf"

case "pps"
icone="pps"

case "wmv"
icone="wmv"

case "wav"
icone="wmv"

case "avi"
icone="avi"

case "mpg"
icone="mpg"

case "mp3"
icone="mpg"

case "mpeg"
icone="mpg"

case "jpg"
icone="jpg"

case "jpeg"
icone="jpg"

case "gif"
icone="gif"

case "bmp"
icone="bmp"

case "rar"
icone="zip"

case "zip"
icone="zip"

end select



data_de=RS_doc("DA_Doc")
if data_de="" or isnull(data_de) then
else			
dados_dtd= split(data_de, "/" )
dia_de= dados_dtd(0)
mes_de= dados_dtd(1)
ano_de= dados_dtd(2)
end if


if dia_de<10 then
dia_de="0"&dia_de
end if
if mes_de<10 then
mes_de="0"&mes_de
end if

data_inicio=dia_de&"/"&mes_de&"/"&ano_de
%>
                        <td width="409" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr valign="top" class="<%response.write(cor)%>"> 
                              <td width="20"><img src="../../img/icones/<%=icone%>.gif" width="18" height="18"></td>
                              <td> <a href="docs/avcirc/download.asp?opt=<%=ordem%>" class="menu_sublista"> 
                                <%response.Write(tit1)%>
                                </a> </td>
                            </tr>
                            <tr valign="top" class="<%response.write(cor)%>"> 
                              <td width="20">&nbsp;</td>
                              <td> 
                                <%response.Write("Publicado em "&data_inicio)
						  RS_doc.movenext
						  %>
                              </td>
                            </tr>
                          </table></td>
                        <td height="40" valign="top"> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr valign="top" class="<%response.write(cor)%>"> 
                              <%
if RS_doc.eof then
%>
                              <td width="20">&nbsp;</td>
                              <td>&nbsp;</td>
                            </tr>
                            <tr valign="top" class="<%response.write(cor)%>"> 
                              <td width="20">&nbsp;</td>
                              <td> 

                              </td>
                            </tr>
                          </table></td>
                      </tr>
<%
else							  
ordem=ordem+1					   
tit1=RS_doc("TI1_Doc")
nome_arq=RS_doc("NO_Doc")
extensao_arq= split(nome_arq, "." )
extensao= extensao_arq(1)
nome_sessao="arq_"&ordem
session(nome_sessao)=nome_arq

select case extensao

case "doc"
icone="word"

case "xls"
icone="excel"

case "pdf"
icone="pdf"

case "pps"
icone="pps"

case "wmv"
icone="wmv"

case "avi"
icone="avi"

case "mpg"
icone="mpg"

case "mpeg"
icone="mpg"

case "jpg"
icone="jpg"

case "jpeg"
icone="jpg"

case "gif"
icone="gif"

end select



data_de=RS_doc("DA_Doc")
if data_de="" or isnull(data_de) then
else			
dados_dtd= split(data_de, "/" )
dia_de= dados_dtd(0)
mes_de= dados_dtd(1)
ano_de= dados_dtd(2)
end if


if dia_de<10 then
dia_de="0"&dia_de
end if
if mes_de<10 then
mes_de="0"&mes_de
end if

data_inicio=dia_de&"/"&mes_de&"/"&ano_de

%>
                              <td width="20"><img src="../../img/icones/<%=icone%>.gif" width="18" height="18"></td>
                              <td> <a href="docs/avcirc/download.asp?opt=<%=ordem%>" class="menu_sublista" > 
                                <%response.Write(tit1)%>
                                </a> </td>
                            </tr>
                            <tr valign="top" class="<%response.write(cor)%>"> 
                              <td width="20">&nbsp;</td>
                              <td> 
                                <%response.Write("Publicado em "&data_inicio)
						  %>
                              </td>
                            </tr>
                          </table></td>
                      </tr>
                      <%
check=check+1
RS_doc.movenext
end if
wend
end if
%>
                    </table></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
  </tr>
  <tr>
    <td width="1000"><img src="img/rodape.jpg" width="1000" height="78" /></td>
  </tr>
</table>
</form>
</body>
</html>
