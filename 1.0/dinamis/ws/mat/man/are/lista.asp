<%'On Error Resume Next%>
<%response.Charset="UTF-8"%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/bd_alunos.asp"-->
<!--#include file="../../../../inc/bd_contato.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<%
pagina = Request.QueryString("pagina")
ordenacao_form = request.form("ordem")

if pagina = "" or isnull(pagina) then
	ordenacao = ordenacao_form
	session("ordenacao") = ordenacao	
	pagina = 1
else
    if isnull(ordenacao_form) or ordenacao_form = "" then
		ordenacao = session("ordenacao")	
	else
		ordenacao = ordenacao_form					
	end if
	session("ordenacao") = ordenacao		
end if	

selected_M = ""
selected_N = ""
selected_UCET = ""
selected_R = ""
selected_D = ""
if ordenacao = "M" then
	selected_M = "selected"
	orderBy = "co_matric"	
elseif ordenacao = "N" then 
	selected_N = "selected"
	orderBy = "nome"	
elseif ordenacao = "UCET" then 
	selected_UCET = "selected"
	orderBy = "unidade,curso,etapa,turma"
elseif ordenacao = "R" then 	
	selected_R = "selected"
	orderBy = "responsavel"		
elseif ordenacao = "D" then 	
	selected_D = "selected"
	orderBy = "data, hora"		
end if

obr = ordenacao

ano_letivo = session("ano_letivo")
co_usr = session("co_user")
nivel=4

nvg = session("nvg")
session("nvg")=nvg

nvg_split=split(nvg,"-")
sistema_local=nvg_split(0)
modulo=nvg_split(1)
setor=nvg_split(2)
funcao=nvg_split(3)

ano_info=nivel&"-"&nvg&"-"&ano_letivo


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CONwf = Server.CreateObject("ADODB.Connection") 
		ABRIRwf = "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONwf.Open ABRIRwf



 call navegacao (CON,nvg,nivel)
navega=Session("caminho")	


%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html;charset=ISO-8859-1">
<script language="JavaScript" src="file:../../../../img/mm_menu.js"></script>
<script type="text/javascript" src="../../../../js/global.js"></script>
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
function submitfuncao()  
{
   var f=document.forms[3]; 
      f.submit(); 
	  
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
//-->
</script>
<script>
function createXMLHTTP()
            {
                        try
                        {
                                   ajax = new ActiveXObject("Microsoft.XMLHTTP");
                        }
                        catch(e)
                        {
                                   try
                                   {
                                               ajax = new ActiveXObject("Msxml2.XMLHTTP");
                                               alert(ajax);
                                   }
                                   catch(ex)
                                   {
                                               try
                                               {
                                                           ajax = new XMLHttpRequest();
                                               }
                                               catch(exc)
                                               {
                                                            alert("Esse browser não tem recursos para uso do Ajax");
                                                            ajax = null;
                                               }
                                   }
                                   return ajax;
                        }
           
           
               var arrSignatures = ["MSXML2.XMLHTTP.5.0", "MSXML2.XMLHTTP.4.0",
               "MSXML2.XMLHTTP.3.0", "MSXML2.XMLHTTP",
               "Microsoft.XMLHTTP"];
               for (var i=0; i < arrSignatures.length; i++) {
                                                                          try {
                                                                                                             var oRequest = new ActiveXObject(arrSignatures[i]);
                                                                                                             return oRequest;
                                                                          } catch (oError) {
                                                                          }
                                      }
           
                                      throw new Error("MSXML is not installed on your system.");
                        }                                
						
						
						 function recuperarCurso(uTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=c", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                    var resultado_c  = oHTTPRequest.responseText;
resultado_c = resultado_c.replace(/\+/g," ")
resultado_c = unescape(resultado_c)
document.all.divCurso.innerHTML =resultado_c
document.all.divEtapa.innerHTML ="<select name='etapa' class='select_style' id='etapa'><option value='999990' selected>           </option></select>"
document.all.divTurma.innerHTML = "<select name='turma' class='select_style' id='turma'><option value='999990' selected>           </option></select>"
document.all.divPeriodo.innerHTML = "<select name='periodo' class='select_style' id='periodo'><option value='0' selected>           </option></select>"
//recuperarEtapa()
                                                           }
                                               }

                                               oHTTPRequest.send("u_pub=" + uTipo);
                                   }


						 function recuperarEtapa(cTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=e", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                      var resultado_e= oHTTPRequest.responseText;
																	   
resultado_e = resultado_e.replace(/\+/g," ")
resultado_e = unescape(resultado_e)
document.all.divEtapa.innerHTML =resultado_e
document.all.divTurma.innerHTML = "<select name='turma' class='select_style' id='turma'><option value='999990' selected>           </option></select>"
document.all.divPeriodo.innerHTML = "<select name='periodo' class='select_style' id='periodo'><option value='0' selected>           </option></select>"
//recuperarTurma()
                                                           }
                                               }

                                               oHTTPRequest.send("c_pub=" + cTipo);
                                   }


						 function recuperarTurma(eTipo)
                                   {

                                               var oHTTPRequest = createXMLHTTP();

                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t4", true);

                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");

                                               oHTTPRequest.onreadystatechange=function() {

                                                           if (oHTTPRequest.readyState==4){

                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t
document.all.divPeriodo.innerHTML = "<select name='periodo' class='select_style' id='periodo'><option value='0' selected>           </option></select>"
																	   
                                                           }
                                               }

                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
								   
function recuperarPeriodo(eTipo)
                                   {
 
                                               var oHTTPRequest = createXMLHTTP();
 
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=p1", true);
 
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
 
                                               oHTTPRequest.onreadystatechange=function() {
 
                                                           if (oHTTPRequest.readyState==4){
 
                                                                       var resultado_p= oHTTPRequest.responseText;
resultado_p = resultado_p.replace(/\+/g," ")
resultado_p = unescape(resultado_p)
document.all.divPeriodo.innerHTML = resultado_p
																	   
                                                           }
                                               }
 
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }									   

 function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}

function checksubmit()
{
  if (document.busca.unidade.value == "999990" || document.busca.curso.value == "999990" || document.busca.etapa.value == "999990" || document.busca.turma.value == "999990" || document.busca.periodo.value == "0")
  { alert("É necessário preencher pelo menos Unidade, Curso, Etapa, Turma e Periodo!")
	var combo = document.getElementById("unidade");
	combo.options[0].selected = "true";
	var combo2 = document.getElementById("curso");
	combo2.options[0].selected = "true";	
	var combo3 = document.getElementById("etapa");
	combo3.options[0].selected = "true";	
	var combo4 = document.getElementById("turma");
	combo4.options[0].selected = "true";	
	var combo5 = document.getElementById("periodo");
	combo5.options[0].selected = "true";		
    return false
  }  
   
  return true
}
</script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<% call cabecalho (nivel)
	  %>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
                    
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> 
    </td>
  </tr>

                  <tr> 
                    
    <td height="10"> 
      <%	call mensagens(4,9714,0,0) 
	  
	  

%>
</td></tr>
<tr>

            <td valign="top"> <form name="busca" method="post" action="lista.asp">
                
        <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Op&ccedil;&otilde;es</td>
          </tr>
          <tr> 
            <td><table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="250" class="tb_subtit">&nbsp;</td>
                  <td width="250" class="tb_subtit">&nbsp;</td>
                  <td width="250" class="tb_subtit"><div align="center">ORDENA&Ccedil;&Atilde;O</div></td>
                  <td width="250" class="tb_subtit">&nbsp;</td>
                  <td width="250" class="tb_subtit">&nbsp;</td>
                </tr>
                <% 'if RS1.EOF THEN %>
                <%'else%>
                <tr> 
                  <td width="250">&nbsp;</td>
                  <td width="250">&nbsp;</td>
                  <td width="250"><div align="center">
                      <select name="ordem" class="select_style" id="ordem">
                        <option value="M" <%response.Write(selected_M)%>>Matricula</option>   
                        <option value="N" <%response.Write(selected_N)%>>Nome do Aluno</option>  
                        <option value="UCET" <%response.Write(selected_UCET)%>>Unidade, Curso, Etapa e Turma</option>
                        <option value="R" <%response.Write(selected_R)%>>Nome do Respons&aacute;vel</option>       
                        <option value="D" <%response.Write(selected_D)%>>Data da Rematr&iacute;cula</option>                        
                      </select>
                  </div></td>
                  <td width="250">&nbsp;</td>
                  <td width="250">&nbsp;</td>
                </tr>
                <%'end if %>
                <tr> 
                  <td height="15" colspan="5" bgcolor="#FFFFFF"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td colspan="3"><hr></td>
                      </tr>
                    <tr>
                      <td width="33%">&nbsp;</td>
                      <td width="34%">&nbsp;</td>
                      <td width="33%"><div align="center"><font size="3" face="Courier New, Courier, mono">
                        <input type="submit" name="Submit2" value="Prosseguir" class="botao_prosseguir">
                      </font></div></td>
                    </tr>
                  </table></td>
                </tr>
              </table></td>
          </tr><tr><td><hr></td></tr>
                <tr class="tb_tit"><td>Alunos Rematriculados</td></tr>
          <tr>
                  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr class="tb_subtit">
                      <td align="center">Matricula</td>
                      <td align="center">Nome do Aluno</td>
                      <td align="center">Nome do Respons&aacute;vel Financeiro</td>
                      <td align="center">CPF</td>
                      <td align="center">Data da Rematr&iacute;cula.&nbsp;</td>
                    </tr>
                    <%		
					
Set Rs_ordena = Server.CreateObject ( "ADODB.RecordSet" )
'Vamos adicionar 2 campos nesse recordset!
'O método Append recebe 3 parâmetros:
'Nome do campo, Tipo, Tamanho (opcional)
'O tipo pertence à um DataTypeEnum, e você pode conferir os tipos em
'http://msdn.microsoft.com/library/default.asp?url=/library/en-us/ado270/htm/mdcstdatatypeenum.asp
'200 -> VarChar (String), 7 -> Data, 139 -> Numeric
Rs_ordena.Fields.Append "co_matric", 139, 10
Rs_ordena.Fields.Append "nome", 200, 255
Rs_ordena.Fields.Append "responsavel", 200, 255
Rs_ordena.Fields.Append "cpf", 200, 255
Rs_ordena.Fields.Append "data", 7
Rs_ordena.Fields.Append "hora", 7
Rs_ordena.Fields.Append "unidade", 200, 255
Rs_ordena.Fields.Append "curso", 200, 255
Rs_ordena.Fields.Append "etapa", 200, 255
Rs_ordena.Fields.Append "turma", 200, 255

'Vamos abrir o Recordset!
Rs_ordena.Open					
					Set RS = Server.CreateObject("ADODB.Recordset")
					SQL = "SELECT * FROM TB_Aunos_Rematriculados"
					RS.Open SQL, CONwf, 3, 3
	gera_link = "N"
	if RS.EOF then
%>
 					<tr class="tb_fundo_linha_par">
                      <td height="50" colspan="5" align="center">Nenhum aluno rematriculado</td>
                    </tr>

<%  else

		While Not RS.EOF 
			matric = RS("CO_Matricula_Escola")
			data = RS("DA_Ult_Acesso") 
			hora = RS("HO_ult_Acesso") 			
			aluno = buscaAluno(matric)
		    vetorAluno = split(aluno,"#!#")
			nome = Server.HTMLEncode(vetorAluno(2))
			tipo_resp_fin = buscaTipoResponsavelFinanceiro(matric)
			
			vetorContato = buscaContato (matric, tipo_resp_fin)
			dadosContato = split(vetorContato, "#!#")
			contratante = Server.HTMLEncode(dadosContato(2))
			cpfContratante = dadosContato(4)	
			
			ucet = buscaUCET(matric,session("ano_letivo"))
			vetorUCET = split(ucet,"#!#")
			nu_unidade =  vetorUCET(0)
			co_curso = vetorUCET(1)
			co_etapa = vetorUCET(2)
			co_turma = vetorUCET(3)
			
			Rs_ordena.AddNew
			Rs_ordena.Fields("co_matric").Value = matric
			Rs_ordena.Fields("nome").Value = nome
			Rs_ordena.Fields("responsavel").Value = contratante	
			Rs_ordena.Fields("cpf").Value = cpfContratante						
			Rs_ordena.Fields("data").Value = data
			Rs_ordena.Fields("hora").Value = hora
			Rs_ordena.Fields("unidade").Value = nu_unidade
			Rs_ordena.Fields("curso").Value = co_curso
			Rs_ordena.Fields("etapa").Value = co_etapa
			Rs_ordena.Fields("turma").Value = co_turma					
			'RS.AddNew
			
		RS.movenext
		wend	
		
		Rs_ordena.Sort = orderBy	

Rs_ordena.PageSize = 30

if Rs_ordena.PageCount>1 then
	gera_link = "S"
end if

'Aqui definimos a página atual
if pagina ="" then
      intpagina = 1
	  Rs_ordena.MoveFirst
else
    if cint(pagina)<1 then
	intpagina = 1
    else
		if cint(pagina)>Rs_ordena.PageCount then  
	    intpagina = Rs_ordena.PageCount
        else
    	intpagina = pagina
		end if
    end if   
 end if   

    Rs_ordena.AbsolutePage = intpagina



Dim RowCount 
RowCount = 0
check=2
					While Not Rs_ordena.EOF And RowCount < Rs_ordena.PageSize
						if check mod 2 =0 then
					  		cor = "tb_fundo_linha_par" 
					 	else cor ="tb_fundo_linha_impar"
					  	end if
%>
                    <tr class="<%response.Write(cor)%>">
                      <td align="center"><%response.write(Rs_ordena.Fields("co_matric").Value)%></td>
                      <td align="center"><%response.write(Rs_ordena.Fields("nome").Value)%></td>
                      <td align="center"><%response.write(Rs_ordena.Fields("responsavel").Value)%></td>
                      <td align="center"><%response.write(Rs_ordena.Fields("cpf").Value)%></td>
                      <td align="center"><%response.write(Rs_ordena.Fields("data").Value&" "&Rs_ordena.Fields("hora").Value)%></td>
                    </tr>
                    
                    <%
					   RowCount = RowCount + 1
     
			Rs_ordena.MoveNext
			Wend
end if
					%>
                     <tr class="tb_tit">
                      <td colspan="5" align="center">
					  <% if gera_link = "S" then	  
			   				if intpagina>1 then
							%>
                            <a href="lista.asp?pagina=<%=intpagina-1%>" class="linktres">Anterior</a> 
                            <%
							end if 
							for contapagina=1 to Rs_ordena.PageCount 
								pagina=pagina*1
								IF contapagina=pagina then
								response.Write(contapagina)
								else
								%>
								<a href="lista.asp?pagina=<%=contapagina%>" class="linktres"><%response.Write(contapagina)%></a> 
								<%
								end if
							next
							if StrComp(intpagina,Rs_ordena.PageCount)<>0 then  
							%>
                                <a href="lista.asp?pagina=<%=intpagina + 1%>" class="linktres">Pr&oacute;ximo</a> 
                                <%
							end if					  
					  	 end if
					  %>
					  </td>
                    </tr>

            </table></td>
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
'response.redirect("../../../../inc/erro.asp")
end if
%>