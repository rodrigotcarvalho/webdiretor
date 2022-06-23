<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes4.asp"-->


<!--#include file="../../../../inc/caminhos.asp"-->





<%opt = request.QueryString("opt")
nvg = session("chave")
co_usr = session("co_user")
chave=nvg
session("chave")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
nivel=4
nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)



if opt = "err1" OR opt = "err2" OR opt = "err3"then
dd=request.querystring("dd")
dados = split(dd,"_")
ano_letivo = dados(0)
curso = dados(1)
unidade = dados(2)
co_grupo = dados(3)
else

co_grupo = request.Form("co_grupo")
curso = request.Form("curso")
unidade = request.Form("unidade")
co_etapa= request.form("etapa")
turma= request.form("turma")
end if
obr=unidade&"_"&curso&"_"&co_etapa&"_"&turma&"_"&periodo&"_"&ano_letivo

minimo_recuperacao=50

nivel=4



		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
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

 call navegacao (CON,chave,nivel)
navega=Session("caminho")

grava_nota= unidade&"-"&curso&"-"&co_etapa&"-"&turma
session("grava_nota")=grava_nota

	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
<script type="text/javascript" src="../../../../js/global.js"></script>
<script language="JavaScript">
 window.history.forward(1);
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
function submitforminterno()  
{
   var f=document.forms[3]; 
      f.submit(); 
	  
} function checksubmit()
{
  if (document.inclusao.etapa.value == "999999")
  {    alert("Por favor, selecione uma etapa!")
    document.inclusao.etapa.focus()
    return false
  }         	     
  return true
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
<body link="#6699CC" vlink="#6699CC" alink="#6699CC" leftmargin="0" background="../../../../img/fundo.gif" topmargin="0" marginwidth="0" marginheight="0">
<% call cabecalho (nivel) %>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF" bgimage="../../../../fundo_interno.gif">
  <tr>     
    <td height="10" valign="top" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)
%></font> </td>
  </tr>
      <%if opt = "err1" then%>
  <tr> 
    <td height="10" valign="top"> 
      <%
		call mensagens(nivel,231,1,0)
%>
    </td>
                </tr>
                <%end if%>
                <%if opt = "err2" then%>
                <tr> 
                  
    <td height="10" valign="top"> 
      <%
		call mensagens(nivel,232,1,0)
%>
    </td>
                </tr>
                <%end if%>
                <%if opt = "err3" then%>
                <tr> 
                  
    <td height="10" valign="top"> 
      <%
		call mensagens(nivel,233,1,0)
%>
    </td>
                </tr>
                <%end if%>
<tr height="10">                  
    <td height="10"><%	call mensagens(nivel,636,0,0) %></td>
                </tr>
<tr>
            <td valign="top"> 
<form name="inclusao" method="post" action="notas.asp" onSubmit="return checksubmit()">
                
        <table width="1000" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td height="15" class="tb_tit">Grade de Aulas
<input name="co_grupo" type="hidden" id="co_grupo" value="<% = co_grupo %>"></td>
          </tr>
          <tr> 
            <td valign="top">
<table width="1000" border="0" cellpadding="0" cellspacing="0">
                <tr> 
                  <td width="12"><font size="1">&nbsp;</font></td>
                  <td width="247" class="tb_subtit"> <div align="center">UNIDADE 
                    </div></td>
                  <td width="247" class="tb_subtit"> <div align="center">CURSO 
                    </div></td>
                  <td width="247" class="tb_subtit"> <div align="center">ETAPA 
                    </div></td>
                  <td width="247" class="tb_subtit"> <div align="center">TURMA 
                    </div></td>
                </tr>
                <tr> 
                  <td width="12"> </td>
                  <td width="247"> <div align="center"> <font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <%response.Write(no_unidade)%>
                      <input name="unidade" type="hidden" id="unidade" value="<% = unidade %>">
                      </font></div></td>
                  <td width="247"> <div align="center"> <font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <%
response.Write(no_curso)%>
                      <input type="hidden" name="curso" value="<% = curso %>">
                      </font></div></td>
                  <td width="247"> <div align="center"> 
                      <input name="etapa" type="hidden" id="etapa" value="<%=etapa%>">
                      <font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& co_etapa &"' and CO_Curso ='"& curso &"'" 
		RS3.Open SQL3, CON0
no_etapa=RS3("NO_Etapa")
response.Write(no_etapa)%>
                      </font></div></td>
                  <td width="247"> <div align="center"> 
                      <input name="turma" type="hidden" id="turma" value="<%=turma%>">
                      <font color="#6699CC" size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
                      <%
response.Write(turma)%>
                      </font></div></td>
                </tr>
                <tr> 
                  <td></td>
                  <td colspan="4">&nbsp;</td>
                </tr>
                <tr>
                  <td></td>
                  <td colspan="4"><%
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
                    <table width="1000" border="0" align="right" cellspacing="0">
                      <tr> 
                        <td width="17" class="tb_subtit"> <div align="right">N&ordm;</div></td>
                        <td width="larg" class="tb_subtit"> <div align="center">Nome</div></td>
                        <%For k=0 To ubound(nome_nota)%>
                        <td width="40"class="tb_subtit"> <div align="center"> 
        <% response.Write(nome_nota(k))%>
      </div></td>
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
  cor = "tb_fundo_linha_par" 
 else cor ="tb_fundo_linha_impar"
  end if
  
if nu_chamada=nu_chamada_check then
nu_chamada_check=nu_chamada_check+1%>
                      <tr class="<%=cor%>"> 
                        <td rowspan="2" width="17"> <div align="center"><font class='form_dado_texto'> 
        <%response.Write(nu_chamada)%>
        </font></div></td>
                        <td rowspan="2" width="200"> <div align="left"><font class='form_dado_texto'> 
        <%response.Write(NO_Aluno)%>
        </font></div></td>
                        <%For k=0 To ubound(nome_nota)
%>
                        <td> <div align="center"><font class='form_dado_texto'><%tipo=tipo_mae(k)
			  ordem2=ordem_mae(k)
			  materia=nome_nota(k)
			 'response.Write(">>"&tipo&"-")
			  call Calc_Med_An_Fin(nu_matricula,unidade,curso,co_etapa,turma,materia,ordem2,tipo)

mamp=session("medFin")

If Not IsArray(medias) Then
medias = Array()
End if
ReDim preserve medias(UBound(medias)+1)
medias(Ubound(medias)) = mamp

response.Write(mamp)%>
                            </font></div></td>
                        <%
NEXT%>
                      </tr>
                      <tr class="<%=cor%>"> 
                        <%For k=0 To ubound(nome_nota)
%>
                        <td> <div align="center"><font class='form_dado_texto'>
        <%
mamp=medias(k)
mamp=mamp*1
minimo_recuperacao=minimo_recuperacao*1
If mamp >= minimo_recuperacao then
res="APR"
else
res="REP"
END IF		  
response.Write(res)%>
        </font></div></td>
                        <%NEXT%>
                      </tr>
                      <% 
else
While nu_chamada>nu_chamada_check
%>
                      <tr class="tb_fundo_linha_falta"> 
                        <td width="17" > <div align="center"> 
                            <%response.Write(nu_chamada_check)%>
                            </div></td>
                        <td width="200"> </td>
                        <%For k=0 To ubound(nome_nota)%>
                        <td> </td>
                        <%

NEXT
%>
                      </tr>
                      <%
nu_chamada_check=nu_chamada_check+1	 
wend	
%>
                      <tr class="<%=cor%>"> 
                        <td rowspan="2" width="17"> <div align="center"><font class='form_dado_texto'> 
        <%response.Write(nu_chamada)%>
        </font></div></td>
                        <td rowspan="2" width="200"> <div align="left"><font class='form_dado_texto'> 
        <%response.Write(NO_Aluno)%>
        </font></div></td>
                        <%For k=0 To ubound(nome_nota)
%>
                        <td> <div align="center"><font class='form_dado_texto'>
                            <%tipo=tipo_mae(k)
			  ordem2=ordem_mae(k)
			  materia=nome_nota(k)
			 'response.Write(">>"&ordem2&"-")
			  call Calc_Med_An_Fin(nu_matricula,unidade,curso,co_etapa,turma,materia,ordem2,tipo)

mamp=session("medFin")

If Not IsArray(medias) Then
medias = Array()
End if
ReDim preserve medias(UBound(medias)+1)
medias(Ubound(medias)) = mamp

response.Write(mamp)%>
                            </font></div></td>
                        <%
NEXT%>
                      </tr>
                      <tr class="<%=cor%>"> 
                        <%For k=0 To ubound(nome_nota)
%>
                        <td> <div align="center"><font class='form_dado_texto'>
        <%
mamp=medias(k)
'response.Write(mamp)
mamp=mamp*1
minimo_recuperacao=minimo_recuperacao*1

If mamp >= minimo_recuperacao then
res="APR"
else
res="REP"
END IF		  
response.Write(res)%>
        </font></div></td>
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
            </td>
          </tr>
        </table>
              </form></td>
          </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</body>
<scripheight="10 type="text/javascript">
<!--
  initInputHighlightScript();
//-->
</script>
<%

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
response.redirect("../../../../inc/erro.asp")
end if
%>