<%'On Error Resume Next%>
<!--#include file="inc/caminhos.asp"-->
<!--#include file="inc/funcoes.asp"-->
<%
nivel=0
permissao = session("permissao")
ano_letivo = session("ano_letivo") 
ano_info=nivel&"_0_"&ano_letivo
chave="WR"
session("sistema_local")="WR"

opt=request.QueryString("opt")
pagina=request.QueryString("pagina")

co_usr=session("co_user")

'if pagina="" or isnull(pagina) then
'	pagina=1
'end if	

'if (pagina=1 or pagina="1") and volta="n" then
'
'end if

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set CON_G = Server.CreateObject("ADODB.Connection") 
		ABRIR_G = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_G.Open ABRIR_G		

		Set CON_M = Server.CreateObject("ADODB.Connection") 
		ABRIR_M = "DBQ="& CAMINHO_msg & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_M.Open ABRIR_M
		
		Set CON_P = Server.CreateObject("ADODB.Connection") 
		ABRIR_P = "DBQ="& CAMINHO_p & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_P.Open ABRIR_P		
		
call navegacao (CON,chave,nivel)
navega=Session("caminho")

exibe_msg="s"

%>
<html>
<head>
<title>Web Diretor</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="JavaScript" src="js/global.js"></script>

<script language="JavaScript" type="text/JavaScript">
<!--

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
      img.src = (img.MM_dn && args[i+2]) ? args[i+2] : ((args[i+1])?args[i+1] : img.MM_up);
      nbArr[nbArr.length] = img;
    }
  } else if (event == "out" ) {
    for (i=0; i < document.MM_nbOver.length; i++) { img = document.MM_nbOver[i]; img.src = (img.MM_dn) ? img.MM_dn : img.MM_up; }
  } else if (event == "down") {
    nbArr = document[grpName];
    if (nbArr) for (i=0; i < nbArr.length; i++) { img=nbArr[i]; img.src = img.MM_up; img.MM_dn = 0; }
    document[grpName] = nbArr = new Array();
    for (i=2; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = img.MM_dn = (args[i+1])? args[i+1] : img.MM_up;
      nbArr[nbArr.length] = img;
  } }
}

function MM_preloadImages() { //v3.0
 var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
   var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
   if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}//-->
</script>
<link href="estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="img/fundo.gif" onLoad="<%response.Write(SESSION("onLoad"))%>">
<%
call cabecalho(nivel)
%>
        
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr class="celula" height="10">
                <td><hr>			 
                </td>
            </tr>
          <tr valign="top" bgcolor="#FFFFFF"> 
            <td background="img/fundo_interno.gif">
<%
ano = DatePart("yyyy", now) 
mes = DatePart("m", now) 
dia = DatePart("d", now) 

data_ate=mes&"/"&dia&"/"&ano

	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "Select * from TB_Mensagens Where ((CO_Usuario = "&co_usr&") or (CO_Usuario IS NULL)) and (NT_DT_Pb BETWEEN #01/01/1900# AND #"&data_ate&"#) and ((NT_DT_Vg NOT BETWEEN #01/01/1900# AND #"&data_ate&"#) or NT_DT_Vg is NULL) order by NT_DT_Pb DESC"
'	SQL = "Select * from TB_Mensagens Where ((CO_Usuario = "&co_usr&") or (CO_Usuario IS NULL)) and (NT_DT_Pb BETWEEN #01/01/1900# AND #"&data_ate&"#) order by NT_DT_Pb DESC"
'	SQL = "Select * from TB_Mensagens Where ((CO_Usuario = "&co_usr&") or (CO_Usuario IS NULL)) and (NT_DT_Pb BETWEEN #01/01/1900# AND #"&data_ate&"#) and (NT_DT_Vg BETWEEN #"&data_ate&"# AND #12/31/3000#) order by NT_DT_Pb DESC"		
	RS.Open SQL, CON_M, 3, 3

	
	Set RS2 = Server.CreateObject("ADODB.Recordset")
	SQL2 = "SELECT CO_Grupo FROM TB_Autoriz_Usuario_Grupo Where CO_Usuario = "&co_usr
	RS2.Open SQL2, CON	

	if RS2.EOF then
			exclui_mensagens_designadas="s"	
	else			
		co_grupo=RS2("CO_Grupo")
		
		if co_grupo="PRO" then
	
			Set RS3 = Server.CreateObject("ADODB.Recordset")
			SQL3 = "SELECT CO_Professor FROM TB_Professor Where CO_Usuario = "&co_usr
			RS3.Open SQL3, CON_p	
			
			if RS3.eof then
				exclui_mensagens_designadas="s"					
			else
				co_professor=RS3("CO_Professor")	
			end if	
		else
			exclui_mensagens_designadas="s"		
		end if	
	End if
	

	exibe_link="s"
	Set Rsv = Server.CreateObject ( "ADODB.RecordSet" )
	Rsv.Fields.Append "tipo_msg", 200, 2
	Rsv.Fields.Append "tit_msg", 200, 255
	Rsv.Fields.Append "cont_msg", 201, 999
	Rsv.Fields.Append "unidade_bd_msg", 200, 10, 64
	Rsv.Fields.Append "curso_bd_msg", 200, 10, 64
	Rsv.Fields.Append "etapa_bd_msg", 200, 10, 64
	Rsv.Fields.Append "turma_bd_msg", 200, 10, 64
	Rsv.Fields.Append "data_pb_msg", 7
	Rsv.Fields.Append "data_vg_msg", 200, 11, 64
	Rsv.Open
	intrec=0						    			
'	WHILE intrec<RS.PageSize and  NOT RS.EOF 	
	WHILE NOT RS.EOF 	
		exibe_msg="s"
					
		tipo_msg=RS("TP_Mensagem")
		tit_msg=RS("NT_Titulo")
		cont_msg=RS("NT_Conteudo")
		unidade_bd_msg=RS("Unidade")
		curso_bd_msg=RS("Curso")
		etapa_bd_msg=RS("Etapa")
		turma_bd_msg=RS("Turma")		
		data_pb_msg=RS("NT_DT_Pb")
		data_vg_msg=RS("NT_DT_Vg")		
		
		if exclui_mensagens_designadas="s"	then		
			if unidade_bd_msg<>"" or curso_bd_msg<>"" or  etapa_bd_msg<>"" or  turma_bd_msg<>"" then
				exibe_msg="n"
			end if
		end if	

		if co_grupo="PRO" then
	
			if unidade_bd_msg="" or isnull(unidade_bd_msg) then
				sem_unidade="s"
			ELSE	
				SQL_U=" NU_Unidade= "&unidade_bd_msg&" and "	
			end if
			
			if curso_bd_msg="" or  isnull(curso_bd_msg) or sem_unidade="s" then
			ELSE	
				SQL_C="CO_Curso= '"&curso_bd_msg&"' and "	
			end if
								
			if etapa_bd_msg="" or  isnull(etapa_bd_msg) or sem_unidade="s" then
			ELSE	
				SQL_E="CO_Etapa= '"&etapa_bd_msg&"'and "		
			end if
			
			if turma_bd_msg="" or isnull(turma_bd_msg) or sem_unidade="s" then
			ELSE	
				SQL_T="CO_Turma= '"&turma_bd_msg&"' and"		
			end if			
	
			Set RS4 = Server.CreateObject("ADODB.Recordset")
			SQL4 = "Select NU_Unidade from TB_Da_Aula Where "&SQL_U&SQL_C&SQL_E&SQL_T&" CO_Professor="&co_professor	
			RS4.Open SQL4, CON_G				
						
			if RS4.EOF then
				exibe_msg="n"					
			end if
'			RESPONSE.Write(exibe_msg&"<br>")
		end if			
			
		if exibe_msg="s" then	
			Rsv.AddNew 
			Rsv.Fields("tipo_msg").Value = tipo_msg
			Rsv.Fields("tit_msg").Value = tit_msg
			Rsv.Fields("cont_msg").Value = cont_msg
			Rsv.Fields("unidade_bd_msg").Value = unidade_bd_msg
			Rsv.Fields("curso_bd_msg").Value = curso_bd_msg		
			Rsv.Fields("etapa_bd_msg").Value = etapa_bd_msg		
			Rsv.Fields("turma_bd_msg").Value = turma_bd_msg		
			Rsv.Fields("data_pb_msg").Value = data_pb_msg		
			Rsv.Fields("data_vg_msg").Value = data_vg_msg	
			intrec=intrec+1															
		End If
		
	RS.MOVENEXT
	WEND		
		
'		IF Rsv.EOF then 
'			Rsv.AddNew 	
'			Rsv.Fields("tipo_msg").Value = 0				
'			Rsv.Fields("tit_msg").Value =  "nenhuma mensagem encontrada!"
'		end if

	Rsv.PageSize =10
	
	if Request.QueryString("pagina")="" then
	  intpagina = 1
	else
		if cint(Request.QueryString("pagina"))<1 then
			intpagina = 1
		else
			if cint(Request.QueryString("pagina"))>Rsv.PageCount then  
				intpagina = Rsv.PageCount
			else
				intpagina = Request.QueryString("pagina")
			end if
		end if   
	end if   	

	if Rsv.EOF then
	else
		Rsv.AbsolutePage = intpagina
		conta_registros=0
		if intpagina=1 then
			Rsv.Movefirst
		end if	
		While conta_registros < Rsv.PageSize and not Rsv.EOF

			tipo_msg=Rsv("tipo_msg")
			tit_msg=Rsv("tit_msg")
			cont_msg=Rsv("cont_msg")
			unidade_bd_msg=Rsv("unidade_bd_msg")
			curso_bd_msg=Rsv("curso_bd_msg")
			etapa_bd_msg=Rsv("etapa_bd_msg")
			turma_bd_msg=Rsv("turma_bd_msg")		
			data_pb_msg=Rsv("data_pb_msg")
			data_vg_msg=Rsv("data_vg_msg")		
			
			if tipo_msg="L" then
				classe="lembrete"
				tipo="Lembrete"			
			elseif tipo_msg="I" then						
				classe="importante"	
				tipo="Importante"					
			elseif tipo_msg="U" then				
				classe="urgente"
				tipo="Urgente"					
			end if			

	%>
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
            <tr class="celula">
                <td height="10"> 
<table width="98%" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr>
    <td width="10%" valign="top" ><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr class="<%response.Write(classe)%>">
    <td><div align="center"><%response.Write(tipo)%></div></td>
  </tr>
</table>
</td>
    <td width="3%" class="celula">&nbsp;
</td>
    <td width="87%" class="celula">
        <%response.Write("("&data_pb_msg&")&nbsp;<strong>"&tit_msg&"</strong> - "&cont_msg)%>
</td>
  </tr>
</table>
		 
                </td>
            </tr>
  <tr class="celula" height="10">
                <td><hr>			 
                </td>
            </tr>
<% 		
	conta_registros=conta_registros+1
	Rsv.MOVENEXT
	WEND
end if	
	
'response.Write(intrec&"<="&Rsv.PageSize )
	
if intpagina = 1 and intrec<=Rsv.PageSize then
	exibe_link="n"
end if	
%>            
  <tr class="celula" height="10">
                <td>&nbsp;</td>			 
            </tr>
                      <%
if exibe_link="s" then
	%>
                <tr> 
                  <td height="20" colspan="10" class="tb_tit"><div align="center"> 

                      &nbsp; 
                    <%		  
			    if intpagina>1 then
    %>
                      <a href="inicio.asp?pagina=<%=intpagina-1%>" class="linktres">Anterior</a> 
                      <%
				end if
		for contapagina=1 to Rsv.PageCount 
						intpagina=intpagina*1
						IF contapagina=intpagina then
						response.Write(contapagina)
						else
						%>
						<a href="inicio.asp?pagina=<%=contapagina%>" class="linktres"><%response.Write(contapagina)%></a> 
						<%
						end if
						next
    if StrComp(intpagina,Rsv.PageCount)<>0 then  
    %>
                      <a href="inicio.asp?pagina=<%=intpagina + 1%>" class="linktres">Próximo</a> 
                      <%
    end if
%>                    </div></td>
                </tr>
<%	
else	
	%>
  <tr class="celula" height="10">
                <td>&nbsp;</td>			 
            </tr>
                      <%
end if	
	Rsv.close
	Set Rsv = Nothing
RS.close
Set RS = Nothing
    %>       
</td>
          </tr>
</table>
</td>
          </tr>
          <tr valign="bottom" bgcolor="#FFFFFF" height="40"> 
            <td>
<div align="center"><img src="img/rodape.jpg" width="1000" height="40"></div></td>
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
pasta=arPath(seleciona)
errfile= pasta&"/"&GetFileName
session("errnumb")=errnumb
session("errdesc")=errdesc
session("errfile") = errfile
response.redirect("inc/erro.asp")
end if
%>