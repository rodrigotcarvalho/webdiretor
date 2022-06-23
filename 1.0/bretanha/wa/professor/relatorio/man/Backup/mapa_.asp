<%On Error Resume Next%>
<!--#include file="../inc/funcoes.asp"-->

<!--#include file="../inc/caminho.asp"-->
<!--#include file="../inc/caminhos.asp"-->




<!--#include file="../inc/funcoes4.asp"-->


<%opt = REQUEST.QueryString("opt")
volta= REQUEST.QueryString("volta")

if opt="direto" then

curso = request.querystring("curso")
unidade = request.querystring("unidade")
grade = request.querystring("grade")
turma = "sem turma"
elseif volta="1" then

curso = request.querystring("curso")
unidade = request.querystring("unidade")
co_etapa = request.querystring("etapa")
turma = request.querystring("turma")
else

curso = request.Form("curso")
unidade = request.Form("unidade")
co_etapa = request.Form("etapa")
turma = request.Form("turma")
end if
ano_letivo = session("ano_letivo")
obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&ano_letivo

if curso="1" then

if co_etapa="1" or co_etapa="2" or co_etapa="3" or co_etapa="4" then

minimo_recuperacao=70

else

minimo_recuperacao=60
end if
else
minimo_recuperacao=60
end if

if co_etapa=999999 then 
response.Redirect("consulta_turma_cp2.asp?opt=err2&or=02&curso="&curso&"&unidade="&unidade&"")
else


id0 = " > <a href='tabelas.asp?or=02&volta=1' class='linkum' target='_self'>Emitir Mapão de Médias Anuais</a>"
id1 = " > <a href='tabelas.asp?or=02&volta=1' class='linkum' target='_self'>Seleciona Unidade</a>"
id2 = " > <a href='tabelas2.asp?opt=vt&or=02&curso="&curso&"&unidade="&unidade&"' class='linkum' target='_self'>Seleciona Etapa</a>"
id3 = " > Consultando"



		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1



		Set CON4 = Server.CreateObject("ADODB.Connection")
		ABRIR4 = "DBQ="& CAMINHOa & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& unidade 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& curso &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Abreviado_Curso")



		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"'" 
		RS3.Open SQL3, CON0
		
if RS3.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RS3("NO_Etapa")
end if

	%>
<html>
<head>
<title>Web Acad&ecirc;mico</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../estilos.css" rel="stylesheet" type="text/css">
<% EscreveFuncaoJavaScriptCurso ( CON0 ) %>
<script type="text/javascript" src="../js/global.js"></script>
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
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);
function submitform()  
{
   var f=document.forms[0]; 
      f.submit(); 
}  
function checksubmit()
{
  if (document.inclusao.etapa.value == "")
  {    alert("Por favor, selecione uma etapa!")
    document.inclusao.etapa.focus()
    return false
  }
  if (document.inclusao.turma.value == "")
  {    alert("Por favor, selecione uma turma!")
    document.inclusao.turma.focus()
return false
}
  if (document.inclusao.mat_prin.value == "0")
  {    alert("Por favor, selecione uma disciplina!")
    document.inclusao.mat_prin.focus()
    return false
  }   
  if (document.inclusao.tabela.value == "")
  {    alert("Por favor, selecione uma tabela!")
    document.inclusao.tabela.focus()
    return false
  }                 	     
  return true
}
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head> 
<body link="#6699CC" vlink="#6699CC" alink="#6699CC" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(1)
%>
<table width="100%" border="0" cellspacing="0">
  <tr>
    <td><div align="center"><table width="1008" border="0">
  <tr> 
    <td><font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="../inicio.asp" target="_parent" class="linkum">Web 
      Acad&ecirc;mico</a> 
      <%

	  response.Write(origem&id0&idL&id1&id2&id3)
%>
      
      </strong> </font></td>
  </tr>
</table>
<br>
<table width="1008" border="0" cellspacing="0">
  <tr> 
    <td width="219" valign="top"> <table width="100%" border="0" cellspacing="0">
        <tr> 
          <td> <%	call mensagens(50,2,0) 

%> </td>
        </tr>
      </table></td>
    <td width="785" valign="top"> <form name="inclusao" method="post" action="grade_cp4i.asp" onSubmit="return checksubmit()">
        <table width="770" height="120" border="0" align="right" cellspacing="0" bgcolor="#F8FAFC">
          <tr bgcolor="#D1E0EF"> 
            <td width="653" height="15" bgcolor="#D1E0EF"><font color="#0000CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Ano 
              Letivo:</strong></font><font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%response.Write(ano_letivo)%>
              </font> </td>
          </tr>
          <tr bgcolor="#D1E0EF"> 
            <td height="15" bgcolor="#FFFFFF"> </td>
          </tr>
          <tr bgcolor="#D1E0EF"> 
            <td height="15" bgcolor="#D1E0EF"><font color="#0000CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
              Segmento</strong></font></td>
          </tr>
          <tr> 
            <td><table width="770" border="0" cellspacing="0">
                <tr> 
                  <td width="8"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp;</font></td>
                  <td bgcolor="#E9F0F8"> <div align="center"><font color="#0000CC"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">UNIDADE 
                      </font></strong></font></div></td>
                  <td bgcolor="#E9F0F8"> <div align="center"><font color="#0000CC"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">CURSO 
                      </font></strong></font></div></td>
                  <td bgcolor="#E9F0F8"> <div align="center"><font color="#0000CC"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">ETAPA 
                      </font></strong></font></div></td>
                  <td bgcolor="#E9F0F8"> <div align="center"><font color="#0000CC"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif">TURMA 
                      </font></strong></font></div></td>
                </tr>
                <tr> 
                  <td width="8"> </td>
                  <td> <div align="center"> <font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <%response.Write(no_unidade)%>
                      <input name="unidade" type="hidden" id="unidade" value="<% = unidade %>">
                      </font></div></td>
                  <td> <div align="center"> <font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <%
response.Write(no_curso)%>
                      <input type="hidden" name="curso" value="<% = curso %>">
                      </font></div></td>
                  <td><div align="center"> <font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <%response.Write(no_etapa)%>
                      <input name="etapa" type="hidden" id="etapa" value="<% = co_etapa %>">
                      </font></div></td>
                  <td> <div align="center"> <font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <font color="#6699CC"> 
                      <%response.Write(turma)%>
                      </font></font><font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <input name="turma" type="hidden" id="turma" value="<% = turma%>">
                      </font><font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      </font></div></td>
                </tr>
              </table></td>
          </tr>
          <tr bgcolor="#FFFFFF"> 
            <td bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
        </table>
      </form></td>
  </tr>
  <tr> 
    <td colspan="2" valign="top"> 
      <%
Set RSNN = Server.CreateObject("ADODB.Recordset")
        CONEXAONN = "Select * from TB_Programa_Aula WHERE CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa&"' order by NU_Ordem_Boletim"
        Set RSNN = CON0.Execute(CONEXAONN)
 
materia_nome_check="vazio"
nome_nota="vazio"
i=0
largura = 0
While not RSNN.eof
materia_nome= RSNN("CO_Materia")
    mae=RSNN("IN_MAE")
    fil=RSNN("IN_FIL")
    in_co=RSNN("IN_CO")
    nu_peso=RSNN("NU_Peso")
	ordem=RSNN("NU_Ordem_Boletim")
'response.Write(materia_nome&"-"&ordem)
if mae=TRUE AND fil=true AND in_co=false AND isnull(nu_peso) then

If Not IsArray(tipo_mae) Then
tipo_mae = Array()
End if
tipo = 4
ReDim preserve tipo_mae(UBound(tipo_mae)+1)
tipo_mae(Ubound(tipo_mae)) = tipo

If Not IsArray(ordem_mae) Then
ordem_mae = Array()
End if
ReDim preserve ordem_mae(UBound(ordem_mae)+1)
ordem_mae(Ubound(ordem_mae)) = ordem

If Not IsArray(nome_nota) Then
nome_nota = Array()
End if
If InStr(Join(nome_nota), materia_nome) = 0 Then
ReDim preserve nome_nota(UBound(nome_nota)+1)
nome_nota(Ubound(nome_nota)) = materia_nome
largura=largura+35

i=i+1

RSNN.movenext
else
RSNN.movenext
end if
 

' sub do anterior
elseif mae=false AND fil=true AND in_co=false then

RSNN.movenext


'MCAL


elseif mae=TRUE AND fil=false AND in_co=true AND isnull(nu_peso) then

If Not IsArray(tipo_mae) Then
tipo_mae = Array()
End if
tipo = 2
ReDim preserve tipo_mae(UBound(tipo_mae)+1)
tipo_mae(Ubound(tipo_mae)) = tipo

If Not IsArray(ordem_mae) Then
ordem_mae = Array()
End if
ReDim preserve ordem_mae(UBound(ordem_mae)+1)
ordem_mae(Ubound(ordem_mae)) = ordem

If Not IsArray(nome_nota) Then
nome_nota = Array()
End if
If InStr(Join(nome_nota), materia_nome) = 0 Then
ReDim preserve nome_nota(UBound(nome_nota)+1)
nome_nota(Ubound(nome_nota)) = materia_nome
largura=largura+35

i=i+1
RSNN.movenext
else
RSNN.movenext
end if


'sub do anterior - MATE 1 E MATE2
elseif mae=false AND fil =false AND in_co=True AND isnull(nu_peso) then

RSNN.movenext

elseif mae=TRUE AND fil=false AND in_co=false AND isnull(nu_peso) then

If Not IsArray(tipo_mae) Then
tipo_mae = Array()
End if
tipo = 3
ReDim preserve tipo_mae(UBound(tipo_mae)+1)
tipo_mae(Ubound(tipo_mae)) = tipo

If Not IsArray(ordem_mae) Then
ordem_mae = Array()
End if
ReDim preserve ordem_mae(UBound(ordem_mae)+1)
ordem_mae(Ubound(ordem_mae)) = ordem

If Not IsArray(nome_nota) Then
nome_nota = Array()
End if
If InStr(Join(nome_nota), materia_nome) = 0 Then
ReDim preserve nome_nota(UBound(nome_nota)+1)
nome_nota(Ubound(nome_nota)) = materia_nome
largura=largura+35

i=i+1


 RSNN.movenext
 else
RSNN.movenext
end if

'se não for nenhum
elseif mae=TRUE AND fil=TRUE AND in_co=false then

If Not IsArray(tipo_mae) Then
tipo_mae = Array()
End if
tipo = 1
ReDim preserve tipo_mae(UBound(tipo_mae)+1)
tipo_mae(Ubound(tipo_mae)) = tipo

If Not IsArray(ordem_mae) Then
ordem_mae = Array()
End if
ReDim preserve ordem_mae(UBound(ordem_mae)+1)
ordem_mae(Ubound(ordem_mae)) = ordem

If Not IsArray(nome_nota) Then
nome_nota = Array()
End if
If InStr(Join(nome_nota), materia_nome) = 0 Then
ReDim preserve nome_nota(UBound(nome_nota)+1)
nome_nota(Ubound(nome_nota)) = materia_nome
largura=largura+35

i=i+1

RSNN.movenext
else
RSNN.movenext
end if
end if
wend
larg=1008-(largura/i)
%>
      <table width="1008" border="0" align="right" cellspacing="0">
        <tr> 
          <td width="17" bordercolor="#E9F0F8" bgcolor="#E9F0F8"> <div align="right"><font color="#0000CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>N&ordm;</strong></font></div></td>
          <td width="larg" bordercolor="#E9F0F8" bgcolor="#E9F0F8"> <div align="center"><font color="#0000CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong>Nome</strong></font></div></td>
          <%For k=0 To ubound(nome_nota)%>
          <td width="40" bordercolor="#E9F0F8" bgcolor="#E9F0F8"> <div align="center"><font color="#0000CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
              <% response.Write(nome_nota(k))%>
              </strong></font></div></td>
          <%
Next%>
        </tr>
        <%  check = 2
nu_chamada_check = 1

	Set RSA = Server.CreateObject("ADODB.Recordset")
	CONEXAOA = "Select * from TB_Aluno_esta_Turma WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& co_etapa &"' AND CO_Turma = '"& turma &"' order by NU_Chamada"
	Set RSA = CON4.Execute(CONEXAOA)
 
 While Not RSA.EOF
nu_matricula = RSA("CO_Matricula")
nu_chamada = RSA("NU_Chamada")
medias = Array()

  		Set RSA2 = Server.CreateObject("ADODB.Recordset")
		CONEXAOA2 = "Select * from TB_Aluno WHERE CO_Matricula = "& nu_matricula
		Set RSA2 = CON4.Execute(CONEXAOA2)
  		NO_Aluno= RSA2("NO_Aluno")

 if check mod 2 =0 then
  cor = "#F8FAFC" 
  else cor ="#F1F5FA"
  end if
  
if nu_chamada=nu_chamada_check then
nu_chamada_check=nu_chamada_check+1%>
        <tr bgcolor=<% = cor %>> 
          <td rowspan="2" width="17"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
              <%response.Write(nu_chamada)%>
              </strong></font></div></td>
          <td rowspan="2" width="200"> <div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
              <%response.Write(NO_Aluno)%>
              </strong></font></div></td>
          <%For k=0 To ubound(nome_nota)
%>
          <td> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1"> 
              <%tipo=tipo_mae(k)
			  ordem2=ordem_mae(k)
			  materia=nome_nota(k)
			 'response.Write(">>"&tipo&"-")
			 
			  call Calc_Med_An_Fin(nu_matricula,unidade,curso,co_etapa,turma,materia,ordem2,tipo)
somamp=session("somampAn")
mamp=session("medAn")

If Not IsArray(medias) Then
medias = Array()
End if
ReDim preserve medias(UBound(medias)+1)
medias(Ubound(medias)) = mamp

response.Write(somamp)%>
              </font></div></td>
          <%
NEXT%>
        </tr>		<tr bgcolor=<% = cor %>>
<%For k=0 To ubound(nome_nota)
%>
<td> <div align="center"><strong><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1"> 
              <%
mamp=medias(k)
mamp=mamp*1
minimo_recuperacao=minimo_recuperacao*1
If mamp >= minimo_recuperacao then
res="APR"
else
res="REC"
END IF		  
response.Write(res)%>
              </font></strong></div>
		</td>
<%NEXT%>
		</tr>
        <% 
else
While nu_chamada>nu_chamada_check
%>
        <tr bgcolor="#E4E4E4"> 
          <td width="17" > <div align="center"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <%response.Write(nu_chamada_check)%>
              </font></strong></div></td>
          <td width="200"> <div align="left"><strong><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              </font></strong></div></td>
          <%For k=0 To ubound(nome_nota)%>
          <td> <div align="center"></div></td>
          <%

NEXT
%>
        </tr>
        <%
nu_chamada_check=nu_chamada_check+1	 
wend	
%>
        <tr bgcolor=<% = cor %>> 
          <td rowspan="2" width="17"> <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
              <%response.Write(nu_chamada)%>
              </strong></font></div></td>
          <td rowspan="2" width="200"> <div align="left"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
              <%response.Write(NO_Aluno)%>
              </strong></font></div></td>
          <%For k=0 To ubound(nome_nota)
%>
          <td> <div align="center"><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1"> 
              <%tipo=tipo_mae(k)
			  ordem2=ordem_mae(k)
			  materia=nome_nota(k)
			 'response.Write(">>"&ordem2&"-")
			  call Calc_Med_An_Fin(nu_matricula,unidade,curso,co_etapa,turma,materia,ordem2,tipo)
somamp=session("somampAn")
mamp=session("medAn")

If Not IsArray(medias) Then
medias = Array()
End if
ReDim preserve medias(UBound(medias)+1)
medias(Ubound(medias)) = mamp

response.Write(somamp)%>
              </font></div></td>
          <%
NEXT%>
        </tr>		<tr bgcolor=<% = cor %>>
<%For k=0 To ubound(nome_nota)
%>
<td> <div align="center"><strong><font face="Verdana, Arial, Helvetica, sans-serif"><font size="1"> 
              <%
mamp=medias(k)
'response.Write(mamp)
mamp=mamp*1
minimo_recuperacao=minimo_recuperacao*1

If mamp >= minimo_recuperacao then
res="APR"
else
res="REC"
END IF		  
response.Write(res)%>
              </font></strong></div>
		</td>
<%NEXT%>
		</tr>
        <%
 nu_chamada_check=nu_chamada_check+1	  
end if

	check = check+1
  RSA.MoveNext
  Wend 
%>
      </table></td>
  </tr>
</table>
<p>&nbsp;</p>
<table width="1008" border="0" cellspacing="0">
  <tr>
    <td width="238">&nbsp;</td>
    <td width="770" bgcolor="#F1F5FA"><font color="#669999" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="../professores.asp" target="_parent" class="voltar1">&lt; 
      Voltar para o menu Professores</a></strong></font></td>
  </tr>
</table>
<p align="center">&nbsp;</p></div></td>
  </tr>
</table>

</body>
<script type="text/javascript">
<!--
  initInputHighlightScript();
//-->
</script>
<%end if 

cod_usr=session("codigo")
call GravaLog (chave,obr)
%>
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
response.redirect("../inc/erro.asp")
end if
%>