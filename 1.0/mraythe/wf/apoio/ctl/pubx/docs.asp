<%'On Error Resume Next%>
<%
opt=request.QueryString("opt")
pagina=request.QueryString("pagina")
volta=request.QueryString("v")

ano_letivo_wf = Session("ano_letivo_wf")
co_usr = session("co_user")
nivel=4
nvg = session("chave")
chave=nvg
session("chave")=chave

exibe="n"

ano_info=nivel&"-"&chave&"-"&ano_letivo_wf

session("tipo_arquivo_upl")="0"

if (pagina=1 or pagina="1") and volta="n" then
tp_doc=request.Form("tipo_doc")
dia_de= request.form("dia_de")
mes_de= request.form("mes_de")
dia_ate=request.Form("dia_ate")
mes_ate=request.Form("mes_ate")
unidade=request.Form("unidade")
curso=request.Form("curso")
etapa=request.Form("etapa")
turma=request.Form("turma")
tit=request.Form("tit")
check_status=request.Form("status")


elseif (pagina=1 or pagina="1") and volta="s"then
dia_de= Session("dia_de")
mes_de= Session("mes_de")
dia_ate=Session("dia_ate")
mes_ate=Session("mes_ate")
unidade=Session("unidade")
curso=Session("curso")
etapa=Session("etapa")
turma=Session("turma")
tit=Session("tit")
check_status=Session("check_status")
tp_doc=session("tipo_arquivo")

else
dia_de= Session("dia_de")
mes_de= Session("mes_de")
dia_ate=Session("dia_ate")
mes_ate=Session("mes_ate")
unidade=Session("unidade")
curso=Session("curso")
etapa=Session("etapa")
turma=Session("turma")
tit=Session("tit")
check_status=Session("check_status")
tp_doc=session("tipo_arquivo")

end if

Session("dia_de")=dia_de
Session("mes_de")=mes_de
Session("dia_ate")=dia_ate
Session("mes_ate")=mes_ate
Session("unidade")=unidade
Session("curso")=curso
Session("etapa")=etapa
Session("turma")=turma
Session("tit")=tit
Session("check_status")=check_status
session("tipo_arquivo") =tp_doc



tipo_arquivo=tp_doc
%>

<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->

<%




if transicao = "S" then
	area="wd"
	site_escola="www.simplynet.com.br/wd/"&ambiente_escola&"/wf/apoio/ctl/pub"
else
	if left(ambiente_escola,5)= "teste" then
		area="wdteste"
		site_escola="www.simplynet.com.br/"&area&"/"&ambiente_escola&"/wf/apoio/ctl/pub"
	else
		area="wd"
		'site_escola="www.webdiretor.com.br/"&ambiente_escola&"/wf/apoio/ctl/pub"
		site_escola="www.simplynet.com.br/"&area&"/"&ambiente_escola&"/wf/apoio/ctl/pub"		
	end if
end if

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF

if unidade="999990" or unidade="" or isnull(unidade) then
sql_un=""
unidade_nome="Todas"
else
sql_un="(Unidade= '"&unidade&"' OR Unidade is null) AND"

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="&unidade
		RS0.Open SQL0, CON0
		
		unidade_nome = RS0("NO_Unidade")
end if

if curso="999990" or curso="" or isnull(curso) then
sql_cu=""
curso_nome="Todos"
else
sql_cu="(Curso='"&curso&"' OR (Curso  is null)) AND"
		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Curso where CO_Curso='"&curso&"'"
		RS0c.Open SQL0c, CON0
		
		curso_nome = RS0c("NO_Curso")
end if

if etapa="999990" or etapa="" or isnull(etapa) then
sql_et=""
etapa_nome="Todas"
else
sql_et="(Etapa='"&etapa&"' OR (Etapa  is null)) AND"

		Set RS0e = Server.CreateObject("ADODB.Recordset")
		SQL0e = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&etapa&"'"
		RS0e.Open SQL0e, CON0
		
		etapa_nome = RS0e("NO_Etapa")
end if

if turma="999990" or turma="" or isnull(turma) then
sql_tu=""
turma_nome="Todas"
else
sql_tu="(Turma='"&turma&"' OR (Turma  is null)) AND "
turma_nome=turma
end if

if tp_doc=0 or  tp_doc="" or isnull(tp_doc) then
sql_tp_doc=""
else
sql_tp_doc="TP_Doc= "&tp_doc&" AND "

end if

if tit="" or isnull(tit) then
sql_tit=""
tit_nome="Todos"
else
sql_tit="(TI1_Doc like '%"&tit&"%') AND"
tit_nome="Contendo a(s) palavra(s): "&tit
end if

data_de=mes_de&"/"&dia_de&"/"&ano_letivo_wf
data_ate=mes_ate&"/"&dia_ate&"/"&ano_letivo_wf
if dia_de<10 then
dia_de="0"&dia_de
end if
if mes_de<10 then
mes_de="0"&mes_de
end if
data_inicio=dia_de&"/"&mes_de&"/"&ano_letivo_wf
if dia_ate<10 then
dia_ate="0"&dia_ate
end if
if mes_ate<10 then
mes_ate="0"&mes_ate
end if
data_fim=dia_ate&"/"&mes_ate&"/"&ano_letivo_wf



'	response.Write "SELECT * FROM TB_Documentos where "&sql_tp_doc&sql_un&sql_cu&sql_et&sql_tu&sql_tit&"(DA_Doc BETWEEN #"&data_de&"# AND #"&data_ate&"#) AND NO_Doc='"&nome_arquivo&"' order by NO_Doc Desc"



	




 call navegacao (CON,chave,nivel)
navega=Session("caminho")	

 Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Autoriz_Usuario_Grupo Where CO_Usuario = "&co_usr
		RS2.Open SQL2, CON
		
if RS2.EOF then

else		
co_grupo=RS2("CO_Grupo")
End if
%>


<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="file:../../../../img/mm_menu.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
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
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
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
	  
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
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
//-->
</script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" >
<% call cabecalho (nivel)
	  %>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
                    
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
	  </td>
	  </tr>
      <%
if opt = "ok" then%>
  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(nivel,59,2,0)
%>
    </td>
                  </tr>
      <%
elseif opt = "ok1" then%>
  <tr> 
                    
    <td height="10"> 
      <%
		call mensagens(nivel,61,2,0)
%>
    </td>
                  </tr>				  
                  <% 	end if 

%>                  <tr> 
                    
    <td height="10"> 
      <%	call mensagens(4,9707,0,0) 
	  
	  
%>
</td></tr>
<tr>

            <td valign="top"> 
			
	
<%if check_status="n" then%>			
<FORM name="formulario" METHOD="POST" ACTION="confirma.asp?opt=f">
<%ELSE%>
<FORM name="formulario" METHOD="POST" ACTION="confirma.asp?opt=d">
<%END if %>
                
        <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Crit&eacute;rios da pesquisa 
              <input name="co_grupo" type="hidden" id="co_grupo" value="<% = co_grupo %>"></td>
          </tr>
          <tr> 
            <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr class="<%response.write(cor)%>"> 
                  <td colspan="9" valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr> 
                        <td colspan="4"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                            <tr> 
                              <td class="tb_subtit"><div align="center"> Nome da Pasta de Documentos</div></td>
                              <td colspan="2" class="tb_subtit"><div align="center">Per&iacute;odo 
                                  da Publica&ccedil;&atilde;o</div></td>
                              <td class="tb_subtit"><div align="center">T&iacute;tulo 
                                  do Documento</div></td>
                              <td width="10%" class="tb_subtit"><div align="center">Status</div></td>
                            </tr>
                            <tr> 
                              <td><div align="center"><font class="form_dado_texto"> 
                                  <%
if tp_doc=0 or  tp_doc="" or isnull(tp_doc) then
	tp_doc_nome= "Todos"
else

	Set RS_doc = Server.CreateObject("ADODB.Recordset")
	SQL_doc = "SELECT * FROM TB_Tipo_Pasta_Doc where CO_Pasta_Doc ="&tp_doc
	RS_doc.Open SQL_doc, CON0

	if RS_doc.eof then
			Response.Write("Tipo de documento"&tp_doc&" não cadastrado")
			response.End()
	else
			tp_doc_nome=RS_doc("NO_Pasta")	
	end if
end if
response.Write(tp_doc_nome)
%>
                                  </font></div></td>
                              <td colspan="2"><div align="center"><font class="form_dado_texto"> 
                                  <%response.Write(data_inicio)%>
                                  at&eacute; 
                                  <%response.Write(data_fim)%>
                                  </font></div></td>
                              <td><div align="center"><font class="form_dado_texto"> 
                                  <%response.Write(tit_nome)%>
                                  </font></div></td>
                              <td><div align="center"><font class="form_dado_texto"> 
                                  <%select case check_status
case "nulo"
status_nome="Todos"
case "s"
status_nome="OK"
case "n"
status_nome="NÃO"
end select
response.Write(status_nome)
%>
                                  </font></div></td>
                            </tr>
                          </table></td>
                      </tr>
                      <tr> 
                        <td>&nbsp;</td>
                        <td colspan="2">&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr> 
                        <td width="250" class="tb_subtit"> <div align="center">UNIDADE 
                          </div></td>
                        <td width="250" class="tb_subtit"> <div align="center">CURSO 
                          </div></td>
                        <td width="250" class="tb_subtit"> <div align="center">ETAPA 
                          </div></td>
                        <td width="250" class="tb_subtit"> <div align="center">TURMA 
                          </div></td>
                      </tr>
                      <tr> 
                        <td width="250"> <div align="center"> <font class="form_dado_texto"> 
                            <%response.Write(unidade_nome)%>
                            </font> </div></td>
                        <td width="250"> <div align="center"> <font class="form_dado_texto"> 
                            <%response.Write(curso_nome)%>
                            </font> </div></td>
                        <td width="250"> <div align="center"> <font class="form_dado_texto"> 
                            <%response.Write(etapa_nome)%>
                            </font> </div></td>
                        <td width="250"> <div align="center"> <font class="form_dado_texto"> 
                            <%response.Write(turma_nome)%>
                            </font> </div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr class="<%response.write(cor)%>"> 
                  <td colspan="9" valign="top" ><hr width="1000"></td>
                </tr>
                <tr class="<%response.write(cor)%>"> 
                  <td colspan="9" valign="top" class="tb_tit">Documentos</td>
                </tr>
                <tr> 
                  <td width="20" class="tb_subtit"> <div align="center"> 
                      <input type="checkbox" name="todos" class="borda" value="" onClick="this.value=check(this.form.doc)">
                    </div></td>
                  <td width="60" class="tb_subtit"> <div align="center">Data</div></td>
                  <td width="300" class="tb_subtit"> 
                    <div align="left">&nbsp; Nome Publicado</div></td>
                  <td width="360" class="tb_subtit"> 
                    <div align="left">Nome do 
                      Arquivo</div></td>
                  <td width="40" class="tb_subtit"> <div align="center">Un</div></td>
                  <td width="40" class="tb_subtit"> <div align="center">Curso 
                    </div></td>
                  <td width="80" class="tb_subtit"> 
                    <div align="center">Etapa</div></td>
                  <td width="60" class="tb_subtit"> 
                    <div align="center">Turma</div></td>
                  <td width="40" class="tb_subtit"> <div align="center">Status</div></td>
                </tr>
                <tr class="<%response.write(cor)%>"> 
                  <td colspan="9"><hr width="1000"></td>
                </tr>
<%' response.Redirect("http://"&site_escola&"/sndocs/learquivo.asp?al="&ano_letivo_wf&"&tp="&tp_doc)
%>			
				
                <%
if ((pagina=1 or pagina="1") and volta="n") or opt="ok1" or opt="f" then
Session("GuardaVetor")= Empty
	Function puxaXML()
		hora = DatePart("h", now) 
		min = DatePart("n", now)
		seg= DatePart("s", now) 
		Set xmlhttp = server.CreateObject("microsoft.XMLHTTP")
		xmlhttp.open "GET","http://"&site_escola&"/sndocs/"&tp_doc&".xml?t="&seg&min&hora,false
		xmlhttp.setrequestheader "ContentType","text/xml"
		xmlhttp.send()
		puxaXML = xmlhttp.responsexml.xml 
	End Function 
	
	dim rootElement
	dim intQtdElementos
	
	dim doc, xsldoc
	set xsldoc=server.createobject("microsoft.xmldom") 
	set doc = server.CreateObject("microsoft.xmldom")
	doc.async = false
	doc.loadxml (puxaXML())
	

	set rootElement  = doc.documentElement
	
	intQtdElementos = rootElement.childNodes.length
	
	vetor_arquivos = Session("GuardaVetor")

	If Not IsArray(vetor_arquivos) Then vetor_arquivos = Array() End if

	for i = 0 to intQtdElementos-1
		nome_arquivo =rootElement.childNodes(i).text
		
		ReDim preserve vetor_arquivos(UBound(vetor_arquivos)+1)
		vetor_arquivos(Ubound(vetor_arquivos )) = nome_arquivo
		Session("GuardaVetor") = vetor_arquivos
	next
else
	vetor_arquivos=Session("GuardaVetor")
	Session("GuardaVetor")=vetor_arquivos
end if

Set Rs_ordena = Server.CreateObject ( "ADODB.RecordSet" )
'Vamos adicionar 2 campos nesse recordset!
'O método Append recebe 3 parâmetros:
'Nome do campo, Tipo, Tamanho (opcional)
'O tipo pertence à um DataTypeEnum, e você pode conferir os tipos em
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/ado270/htm/mdcstdatatypeenum.asp
'200 -> VarChar (String), 7 -> Data, 139 -> Numeric
Rs_ordena.Fields.Append "nome", 200, 500
Rs_ordena.Fields.Append "data", 7
Rs_ordena.Fields.Append "co_doc", 139, 10
Rs_ordena.Fields.Append "tipo_doc", 200, 255
Rs_ordena.Fields.Append "tit1", 200, 255
Rs_ordena.Fields.Append "nome_arq", 200, 255
Rs_ordena.Fields.Append "da_doc", 7
Rs_ordena.Fields.Append "unidade", 200, 255
Rs_ordena.Fields.Append "curso", 200, 255
Rs_ordena.Fields.Append "etapa", 200, 255
Rs_ordena.Fields.Append "turma", 200, 255
Rs_ordena.Fields.Append "status", 200, 255

'Vamos abrir o Recordset!
Rs_ordena.Open

'Temos que percorrer agora todos os arquivos e jogar na nossa tabela virtual!

check=2
conta_arquivos=0
for i = 0 to ubound(vetor_arquivos)
nome_arquivo =vetor_arquivos(i)

if nome_arquivo="GIATHE_-_Gincana_Acampamento_do_Colegio_Maria_Raythe.pdf" then
else

'response.Write("'"&nome_arquivo&"'<BR>")
	if check_status="n" then

		Set RS_doc = Server.CreateObject("ADODB.Recordset")
		SQL_doc = "SELECT * FROM TB_Documentos where "&sql_tp_doc&" NO_Doc='"&nome_arquivo&"' order by DA_Doc,TI1_Doc Desc"
		RS_doc.Open SQL_doc, CON_WF, 3, 3

		IF RS_doc.EOF then
			Rs_ordena.AddNew
			Rs_ordena.Fields("nome").Value = nome_arquivo
			'Rs_ordena.Fields("data").Value = arquivo.DateLastModified
			Rs_ordena.Fields("status").Value = "Não"
			Rs_ordena.Fields("unidade").Value = "nulo"
			Rs_ordena.Fields("curso").Value = "nulo"
			Rs_ordena.Fields("etapa").Value = "nulo"
			Rs_ordena.Fields("turma").Value = "nulo"	
			
			conta_arquivos=conta_arquivos+1		
		END IF			
				
	else 				
		Set RS_doc = Server.CreateObject("ADODB.Recordset")
		SQL_doc = "SELECT * FROM TB_Documentos where "&sql_tp_doc&sql_un&sql_cu&sql_et&sql_tu&sql_tit&"(DA_Doc BETWEEN #"&data_de&"# AND #"&data_ate&"#) AND NO_Doc='"&nome_arquivo&"' order by NO_Doc Desc"
		RS_doc.Open SQL_doc, CON_WF, 3, 3


		while not RS_doc.eof
			co_doc=RS_doc("CO_Doc") 
			tipo_doc =RS_doc("TP_Doc") 
			tit1=RS_doc("TI1_Doc")
			nome_arq=RS_doc("NO_Doc")
			da_doc=RS_doc("DA_Doc")
			unidade=RS_doc("Unidade")
			curso=RS_doc("Curso")
			etapa=RS_doc("Etapa")
			turma=RS_doc("Turma")
			if unidade="" or isnull(unidade) then
				unidade="nulo"
			end if
			
			if curso="" or isnull(curso) then
				curso="nulo"
			end if
			
			if etapa="" or isnull(etapa) then
				etapa="nulo"
			end if
			
			if turma="" or isnull(turma) then
				turma="nulo"
			end if
			
			Rs_ordena.AddNew
			Rs_ordena.Fields("nome").Value = nome_arquivo
			'Rs_ordena.Fields("data").Value = arquivo.DateLastModified
			Rs_ordena.Fields("co_doc").Value = co_doc
			Rs_ordena.Fields("tipo_doc").Value = tipo_doc
			Rs_ordena.Fields("tit1").Value = tit1
			Rs_ordena.Fields("nome_arq").Value = nome_arq
			Rs_ordena.Fields("da_doc").Value = da_doc
			Rs_ordena.Fields("unidade").Value = unidade
			Rs_ordena.Fields("curso").Value = curso
			Rs_ordena.Fields("etapa").Value = etapa
			Rs_ordena.Fields("turma").Value = turma
			Rs_ordena.Fields("status").Value = "OK"
			
		RS_doc.movenext
		wend

	end if
end if

next

'Todos os arquivos no recordset, agora vamos ordená-lo!
'Da maior data para a menor!
Rs_ordena.Sort = "da_doc DESC, nome ASC"
if Rs_ordena.EOF then
	exibe_proximo = "N" 
		%>
                <tr class="tb_fundo_linha_par"> 
                  <td colspan="10" valign="top"> <div align="center"><font class="style1"> 
                      <%response.Write("Não existem documentos para os critérios informados!")%>
                      </font></div></td>
                </tr>
                <%
ELSE				
'Pronto! Agora temos os arquivos todos em ordem em nosso recordset! Vamos exibi-los!
	exibe_proximo = "S" 

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

if Rs_ordena.Fields("unidade").Value ="nulo" then
no_unidade=""
else

	if Rs_ordena.Fields("unidade").Value =999990 then
	no_unidade=""
	else
			Set RSnoun = Server.CreateObject("ADODB.Recordset")
			SQLnoun = "SELECT * FROM TB_Unidade Where NU_Unidade="&Rs_ordena.Fields("unidade").Value 
			RSnoun.Open SQLnoun, CON0
			
		no_unidade=RSnoun("NO_Abr")
	end if
end if		

if Rs_ordena.Fields("curso").Value="nulo" then
no_curso=""
else
 		Set RSnocu = Server.CreateObject("ADODB.Recordset")
		SQLnocu = "SELECT * FROM TB_Curso Where CO_Curso='"&Rs_ordena.Fields("curso").Value&"'"
		RSnocu.Open SQLnocu, CON0
		
no_curso=RSnocu("NO_Abreviado_Curso")		
end if

if Rs_ordena.Fields("etapa").Value="nulo" then
no_etapa=""
else


 		Set RSnoet = Server.CreateObject("ADODB.Recordset")
		SQLnoet = "SELECT * FROM TB_Etapa Where CO_Curso='"&Rs_ordena.Fields("curso").Value&"' AND CO_Etapa='"&Rs_ordena.Fields("etapa").Value&"'"
		RSnoet.Open SQLnoet, CON0
		
no_etapa=RSnoet("NO_Etapa")		
end if

if Rs_ordena.Fields("turma").Value="nulo" then
no_turma=""
else
no_turma=Rs_ordena.Fields("turma").Value
end if

IF check_status="n" THEN
da_show=""
ELSE
data_split= Split(Rs_ordena.Fields("da_doc").Value,"/")
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

da_show=dia&"/"&mes&"/"&ano
END IF
%>
                <tr valign="top" class="<%response.write(cor)%>"> 
                  <td width="20"> 
                    <div align="center"><font class="form_dado_texto"> 
<% if Rs_ordena.Fields("status").Value = "OK" then %>					
                      <input name="doc" type="checkbox" class="borda" value="<%=Rs_ordena.Fields("co_doc").Value%>">
<%else
no_doc_a_alterar=Rs_ordena.Fields("nome").Value
no_doc=replace(no_doc_a_alterar, ", ", "#virgespaco#")
no_doc=replace(no_doc, ",", "#$#")
%>
                      <input name="doc" type="checkbox" class="borda" value="<%=no_doc%>">
<%end if%>					  
                      </font></div></td>
                  <td width="60"> 
                    <div align="center"> 
                      <%response.Write(da_show)%>
                    </div></td>
                  <td width="300"> 
                    <div align="left"> &nbsp;
<% if Rs_ordena.Fields("status").Value = "OK" then %>					
					 <a href="alterar.asp?c=<%=Rs_ordena.Fields("co_doc").Value%>" class="linkum"> 
                      <%response.Write(Rs_ordena.Fields("tit1").Value)%>
                      </a> 
<%end if%>					
</div></td>
                  <td width="360"> 
                    <%response.Write(Rs_ordena.Fields("nome"))%>
                    <div align="left"></div></td>
                  <td width="40"> 
                    <div align="center"> 
                      <%response.Write(no_unidade)%>
                    </div></td>
                  <td width="40"> 
                    <div align="center"> 
                      <%response.Write(no_curso)%>
                    </div></td>
                  <td width="80"> 
                    <div align="center"> 
                      <%response.Write(no_etapa)%>
                    </div></td>
                  <td width="60"> 
                    <div align="center"> 
                      <%response.Write(no_turma)%>
                    </div></td>
                  <td width="40"> 
                    <div align="center"> <%response.Write(Rs_ordena.Fields("status").Value )%> </div></td>
                </tr>
                <%
intrec=intrec+1
check=check+1				
Rs_ordena.movenext
Wend
end if
%>
                <tr>
                  <td colspan="9"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td class="tb_tit"><div align="center">                          </div></td>
                      </tr>
                      <tr> 
                        <td class="tb_tit"><div align="center"> 
                            <%
if sem_link=0 then
	%>
                            &nbsp; 
                            <%		  
			    if intpagina>1 then
    %>
                            <a href="docs.asp?pagina=<%=intpagina-1%>" class="linktres">Anterior</a> 
                            <%
    end if 
	for contapagina=1 to Rs_ordena.PageCount 
						pagina=pagina*1
						IF contapagina=pagina then
						response.Write(contapagina)
						else
						%>
						<a href="docs.asp?pagina=<%=contapagina%>" class="linktres"><%response.Write(contapagina)%></a> 
						<%
						end if
						next
    if StrComp(intpagina,Rs_ordena.PageCount)<>0 and exibe_proximo = "S" then  
    %>
                            <a href="docs.asp?pagina=<%=intpagina + 1%>" class="linktres">Próximo</a> 
                            <%
    end if
else	
	%>
                            &nbsp; 
                            <%
end if	
 Rs_ordena.Close
Set Rs_ordena = Nothing

Set objPasta = Nothing
Set objFSO = Nothing
    %>
                          </div></td>
                      </tr>
                    </table></td>
                </tr>
                <tr> 
                  <td colspan="9"><div align="center"> 
                      <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td colspan="4"><hr></td>
                        </tr>
                        <tr> 
                          <td width="25%"> <div align="center"> 
                              <input name="SUBMIT5" type=button class="botao_cancelar" onClick="MM_goToURL('parent','index.asp?nvg=<%=nvg%>');return document.MM_returnValue" value="Voltar">
                          </div></td>
                          <td width="25%"> <div align="center"> 
                              <input name="SUBMIT3" type=submit class="borda_bot4" value="Excluir">
                            </div></td>
                          <td width="25%"> <div align="center"> 
                          		<%if conta_arquivos>0 then
									ativa_botao=""
								else
									ativa_botao="disabled"
								end if
								%>
                              <input name="SUBMIT2" type=button class="botao_prosseguir" onClick="MM_goToURL('parent','incluir.asp');return document.MM_returnValue" value="Associar" <%response.Write(ativa_botao)%>>
                            </div></td>
                          <td width="25%"> <div align="center"> 
                              <input name="Button" type=button class="botao_prosseguir" onClick="MM_goToURL('parent','upload.asp');return document.MM_returnValue" value="Publicar">
                          </div></td>
                        </tr>
                        <tr> 
                          <td>&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>&nbsp;</td>
                          <td>&nbsp;</td>
                        </tr>
                      </table>
                  </div></td>
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
</html>
<%call GravaLog (nvg,"")%>
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