<%On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->

<!--#include file="../../../../inc/funcoes2.asp"-->

<%
Session.LCID = 1046
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg=request.QueryString("nvg")
opt = request.QueryString("opt")
chave=nvg
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo


if opt="1" then
			call GravaLog (chave,"Ativado")
end if

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CONL = Server.CreateObject("ADODB.Connection") 
		ABRIRL = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source="&CAMINHO_log&";"
		CONL.Open ABRIRL
		
if Request.QueryString("pagina")="" then
      intpagina = 1
ano_mnl = request.form("ano_mnl")
mes_mnl = request.form("mes_mnl")
dia_mnl = request.form("dia_mnl")
hora_mnl = request.form("hora_mnl")
min_mnl = request.form("min_mnl")

data_consulta=mes_mnl&"/"&dia_mnl&"/"&ano_mnl
hora_consulta=hora_mnl&":"&min_mnl&":00"
Session("data_consulta")=data_consulta
Session("hora_consulta")=hora_consulta
		Set RSL= Server.CreateObject("ADODB.Recordset")
SQLL = "SELECT * FROM TB_Log_Ocorrencias WHERE (((TB_Log_Ocorrencias.CO_Modulo)='LN') AND ((TB_Log_Ocorrencias.CO_Funcao)='LAN') AND ((TB_Log_Ocorrencias.DA_Ult_Acesso)=#"&data_consulta&"#) AND ((TB_Log_Ocorrencias.HO_ult_Acesso)>=#12/30/1899 "&hora_consulta&"#)) OR (((TB_Log_Ocorrencias.CO_Modulo)='PF') AND ((TB_Log_Ocorrencias.CO_Funcao)='MNL') AND ((TB_Log_Ocorrencias.DA_Ult_Acesso)=#"&data_consulta&"#) AND ((TB_Log_Ocorrencias.HO_ult_Acesso)>=#12/30/1899 "&hora_consulta&"#))order by DA_Ult_Acesso,HO_ult_Acesso"		
		RSL.Open SQLL, CONL, 3, 3
else
data_consulta=Session("data_consulta")
hora_consulta=Session("hora_consulta")
Session("data_consulta")=data_consulta
Session("hora_consulta")=hora_consulta

data_desc = split(data_consulta,"/")
mes_mnl =data_desc(0)
dia_mnl =data_desc(1)
ano_mnl=data_desc(2)

hora_desc = split(hora_consulta,":")
hora_mnl=data_desc(0)
min_mnl=data_desc(1)


		Set RSL= Server.CreateObject("ADODB.Recordset")
		'SQLL = "SELECT * FROM TB_Log_Ocorrencias WHERE CO_Modulo = 'LANWN' AND CO_Funcao = 'LAN' AND DA_Ult_Acesso >= "&DATESERIAL(ano_mnl, mes_mnl, dia_mnl)&" AND HO_ult_Acesso>= #"&hora_consulta&"# order by DA_Ult_Acesso,HO_ult_Acesso"		
		'SQLL = "SELECT * FROM TB_Log_Ocorrencias WHERE (((TB_Log_Ocorrencias.CO_Modulo)='LANWN') AND ((TB_Log_Ocorrencias.CO_Funcao)='LAN') AND ((TB_Log_Ocorrencias.DA_Ult_Acesso)>= #"&data_consulta&"#) AND ((TB_Log_Ocorrencias.HO_ult_Acesso)>=#12/30/1899 "&hora_consulta&"#)) order by DA_Ult_Acesso,HO_ult_Acesso"
SQLL = "SELECT * FROM TB_Log_Ocorrencias WHERE (((TB_Log_Ocorrencias.CO_Modulo)='LN') AND ((TB_Log_Ocorrencias.CO_Funcao)='LAN') AND ((TB_Log_Ocorrencias.DA_Ult_Acesso)=#"&data_consulta&"#) AND ((TB_Log_Ocorrencias.HO_ult_Acesso)>=#12/30/1899 "&hora_consulta&"#)) OR (((TB_Log_Ocorrencias.CO_Modulo)='PF') AND ((TB_Log_Ocorrencias.CO_Funcao)='MNL') AND ((TB_Log_Ocorrencias.DA_Ult_Acesso)=#"&data_consulta&"#) AND ((TB_Log_Ocorrencias.HO_ult_Acesso)>=#12/30/1899 "&hora_consulta&"#))order by DA_Ult_Acesso,HO_ult_Acesso"		

		RSL.Open SQLL, CONL, 3, 3
		
    if cint(Request.QueryString("pagina"))<1 then
	intpagina = 1
    else
		if cint(Request.QueryString("pagina"))>RSL.PageCount then  
	    intpagina = RSL.PageCount
        else
    	intpagina = Request.QueryString("pagina")
		end if
    end if   
 end if

	
 RSL.PageSize = 30
 
if Request.QueryString("pagina")="" then
      intpagina = 1
else
    if cint(Request.QueryString("pagina"))<1 then
	intpagina = 1
    else
		if cint(Request.QueryString("pagina"))>RSL.PageCount then  
	    intpagina = RSL.PageCount
        else
    	intpagina = Request.QueryString("pagina")
		end if
    end if   
 end if   

 
 call VerificaAcesso (CON,chave,nivel)
autoriza=Session("autoriza")

 call navegacao (CON,chave,nivel)
navega=Session("caminho")


	%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
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
function submitfuncao()  
{
   var f=document.forms[0]; 
      f.submit(); 
	  
}  function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
</head> 
<body background="../../../../img/fundo.gif" alink="#CC9900" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<form name="form1" method="post" action="monitora.asp?or=02" onSubmit="return checksubmit()">
<%call cabecalho(nivel)%>
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">		  				  				  
                  <tr>                    
            
    <td height="10" class="tb_caminho"> <font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
      </font> </td>
  </tr>
   <tr>                   
    <td height="10"> 
      <%
	  if autoriza="no" then
	  	call mensagens(4,9700,1,0) 	  
	  else
	  	call mensagens(4,618,0,0) 
	  end if%>
    </td>
                  </tr> 
  <tr> 
    <td valign="top">

<table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
        <tr> 
          <td> 
            <%	  
		if autoriza="no" then			
		else
ano_slct = DatePart("yyyy", now)
mes_slct = DatePart("m", now) 
dia_slct = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

hora = hora*1
min = min*1		
%>
            <table width="1000" border="0" cellspacing="0">
              <tr> 
                <td valign="top"> 
                  <table width="100%" border="0" align="right" cellspacing="0" class="tb_corpo"
>
                    <tr> 
                      <td class="tb_tit">Monitorando Notas</td>
                    </tr>
                    <tr> 
                      <td>
                          <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td class="tb_subtit"> <div align="center"><strong><strong>DATA</strong></strong></div></td>
                              <td class="tb_subtit"> <div align="center"><strong>HORA</strong></div></td>
                            </tr>
                            <tr> 
                              <td width="43%"><div align="center"> <font class="form_dado_texto"> 
                                  <input name="dia_mnl" type="hidden" id="dia" value="<%=dia_mnl%>">
                                  <input name="mes_mnl" type="hidden" id="mes" value="<%=mes_mnl%>">
                                  <input name="ano_mnl" type="hidden" id="ano" value="<%=ano_mnl%>">
                                  <%response.Write(dia_mnl)%>
                                  / 
                                  <%
					  if mes_mnl< 10 then
					  mes_wrt="0"&mes_mnl
					  else
					  mes_wrt=mes_mnl					  
					  end if 
					  response.Write(mes_wrt)%>
                                  / 
                                  <%response.Write(ano_mnl)%>
                                  </font> </div></td>
                              <td width="57%"><div align="center"> <font class="form_dado_texto"> 
                                  <%				  
					  if min_mnl< 10 then
					  min_wrt="0"&min_mnl
					  else
					  min_wrt=min_mnl
					  end if 
					  
					  %>
                                  <input name="hora_mnl" type="hidden" id="hora" value="<%=hora_mnl%>">
                                  <input name="min_mnl" type="hidden" id="min" value="<%=min_mnl%>">
                                  <%response.Write(hora_mnl)%>
                                  : 
                                  <%response.Write(min_wrt)%>
                                  </font> </div></td>
                            </tr>
                            <tr> 
                              <td height="5" colspan="2"> </td>
                            </tr>
                            <tr> 
                              <td colspan="2"> <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td width="70" class="tb_subtit"> 
                                      <div align="center"><strong>Unidade</strong></div></td>
                                    <td width="70" class="tb_subtit"> 
                                      <div align="center"><strong>Curso</strong></div></td>
                                    <td width="70" class="tb_subtit"> 
                                      <div align="center"><strong>Etapa</strong></div></td>
                                    <td width="70" class="tb_subtit"> 
                                      <div align="center"><strong>Turma</strong></div></td>
                                    <td width="125" class="tb_subtit"> 
                                      <div align="center"><strong>Per&iacute;odo</strong></div></td>
                                    <td width="125" class="tb_subtit">
<div align="center"><strong>Disciplina</strong></div></td>
                                    <td width="280" class="tb_subtit"> 
                                      <div align="center"><strong>Planilha 
                                        modificada por</strong></div></td>
                                    <td width="60" class="tb_subtit"> 
                                      <div align="center"><strong>Dia</strong></div></td>
                                    <td width="60" class="tb_subtit"> 
                                      <div align="center"><strong>Hora</strong></div></td>
                                  </tr>
                                  <% 
IF RSL.EOF then
intpagina=1
sem_link=1
%>
<tr> 
                              <td colspan="9"><div align="center"> <font class="form_dado_texto"> Sem Movimento</font></div></td></tr>
<%
else
sem_link=0
    RSL.AbsolutePage = intpagina
    intrec = 0
While intrec<RSL.PageSize and not RSL.EOF
ln_dt = RSL("DA_Ult_Acesso")
ln_hr = RSL("HO_ult_Acesso")


mnl_ln_dt = split(ln_dt,"/")
mnl_dia = mnl_ln_dt(0)
mnl_m = mnl_ln_dt(1)
mnl_a = mnl_ln_dt(2)

mnl_dia = mnl_dia*1
mnl_m = mnl_m*1
mnl_a = mnl_a*1

mnl_ln_hr = split(ln_hr,":")
mnl_h = mnl_ln_hr(0)
mnl_mn = mnl_ln_hr(1)
mnl_h = mnl_h*1
mnl_mn = mnl_mn*1

					  if mnl_m< 10 then
					  mnl_m_wrt="0"&mnl_m
					  else
					  mnl_m_wrt = mnl_m
					  end if 

					  if mnl_mn< 10 then
					  mnl_mn_wrt="0"&mnl_mn
					  else
					  mnl_mn_wrt = mnl_mn
					  end if 


ln_dt = mnl_dia&"/"&mnl_m_wrt&"/"&mnl_a
ln_hr = mnl_h&":"&mnl_mn_wrt

usr_grv = RSL("CO_Usuario")
desc = RSL("TX_Descricao")


mnl_desc = split(desc,",")

mnl_p = mnl_desc(0)
mnl_p_dado = split(mnl_p,":")
mnl_p_dado_tx = mnl_p_dado(1)

mnl_d = mnl_desc(1)
mnl_d_dado = split(mnl_d,":")
mnl_d_dado_tx = mnl_d_dado(1)

mnl_u = mnl_desc(2)
mnl_u_dado = split(mnl_u,":")
mnl_u_dado_tx = mnl_u_dado(1)

mnl_c = mnl_desc(3)
mnl_c_dado = split(mnl_c,":")
mnl_c_dado_tx = mnl_c_dado(1)

mnl_e = mnl_desc(4)
mnl_e_dado = split(mnl_e,":")
mnl_e_dado_tx = mnl_e_dado(1)

mnl_t = mnl_desc(5)
mnl_t_dado = split(mnl_t,":")
mnl_t_dado_tx = mnl_t_dado(1)

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Unidade where NU_Unidade="& mnl_u_dado_tx 
		RS0.Open SQL0, CON0
		
no_unidade = RS0("NO_Unidade")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Curso where CO_Curso ='"& mnl_c_dado_tx &"'"
		RS1.Open SQL1, CON0
		
no_curso = RS1("NO_Abreviado_Curso")

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Etapa where CO_Etapa ='"& mnl_e_dado_tx &"' AND CO_Curso ='"& mnl_c_dado_tx &"'"
		RS3.Open SQL3, CON0
		
if RS3.EOF THEN
no_etapa="sem etapa"
else
no_etapa=RS3("NO_Etapa")
end if

		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia_Principal='"& mnl_d_dado_tx &"'"
		RS7.Open SQL7, CON0

		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_Materia where CO_Materia='"& mnl_d_dado_tx &"'"
		RS8.Open SQL8, CON0

		no_mat= RS8("NO_Materia")

		if RS7.EOF Then						
		co_mat_fil = co_mat_prin
		no_mat_prin = no_mat
		else		
		co_mat_fil= RS7("CO_Materia")
		end if
no_materia=no_mat

	SQL9 = "select * from TB_Usuario where CO_Usuario = " & usr_grv & ""
	set RS9 = CON.Execute (SQL9)

nom_prof=RS9("NO_Usuario")
%>
                                  <tr> 
                                    <td width="70">
<div align="center"> <font class="form_dado_texto"> 
                                        <% response.Write(no_unidade)%>
                                        </font> </div></td>
                                    <td width="70">
<div align="center"> <font class="form_dado_texto"> 
                                        <% response.Write(no_curso)%>
                                        </font> </div></td>
                                    <td width="70">
<div align="center"> <font class="form_dado_texto"> 
                                        <% response.Write(no_etapa)%>
                                        </font> </div></td>
                                    <td width="70">
<div align="center"> <font class="form_dado_texto"> 
                                        <% response.Write(mnl_t_dado_tx)%>
                                        </font> </div></td>
                                    <td width="125"> 
                                      <div align="center"> <font class="form_dado_texto"> 
                                        <% 
							periodo=mnl_p_dado_tx
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo where NU_Periodo="&periodo
		RS4.Open SQL4, CON0


NO_Periodo= RS4("NO_Periodo")
response.Write(NO_Periodo)%>
                                      </div></td>
                                    <td width="125"> 
                                      <div align="center"> <font class="form_dado_texto"> 
                                        <% response.Write(no_materia)%>
                                        </font> </div></td>
                                    <td width="280"> 
                                      <div align="center"> <font class="form_dado_texto"> 
                                        <% response.Write(nom_prof)%>
                                        </font> </div></td>
                                    <td width="60">
<div align="center"> <font class="form_dado_texto"> 
                                        <% response.Write(ln_dt)%>
                                        </font> </div></td>
                                    <td width="60"> <div align="center"><font class="form_dado_texto"> 
                                        <% response.Write(ln_hr)%>
                                        </font> </div></td>
                                  </tr>
                                  <%
intrec = intrec + 1
RSL.Movenext
Wend
End if

%>
                                </table></td>
                            </tr>
                          </table>
                        </td>
                    </tr>
                  </table>
                </td>
              </tr>
            </table></div>
            <%end if 
%>
          </td>
        </tr>
        <tr>
          <td><div align="center">
		  <%for i=1 to RSL.PageCount
		  intpagina=intpagina*1
			  if i=intpagina then%>
			  <font class="form_dado_texto"><%response.Write(intpagina)%></font>
			  <%else%>
			   <a href="monitora.asp?pagina=<%=response.Write(i)%>&nvg=<%=nvg%>&p=<%=periodo%>&c=<%=cod_cons%>&opt=pg" class="linkPaginacao"><%response.Write(i)%></a> 
			  <%end if
		  next
		  %></div></td>
        </tr>
        <tr> 
          <td class="tb_tit"><div align="center">
              <%
if sem_link=0 then
	%>&nbsp;<%		  
			    if intpagina>1 then
    %>
              <a href="monitora.asp?pagina=<%=intpagina-1%>&nvg=<%=nvg%>" class="linktres">Anterior</a> 
              <%
    end if
    if StrComp(intpagina,RSL.PageCount)<>0 then  
    %>
              <a href="monitora.asp?pagina=<%=intpagina + 1%>&nvg=<%=nvg%>" class="linktres">Próximo</a> 
              <%
    end if
else	
	%>&nbsp;<%
end if	
    RSL.close
    Set RSL = Nothing
    %>
            </div></td>
        </tr>
      </table>        
      </td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>
</form>
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