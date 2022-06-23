<%'On Error Resume Next%>
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->



<!--#include file="../../../../inc/caminhos.asp"-->


<% 
opt= request.QueryString("opt")
nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
ori = request.QueryString("or")
chave=session("nvg")
session("nvg")=chave
ano_info=nivel&"-"&chave&"-"&ano_letivo
cod= request.QueryString("cod_cons")	

z = request.QueryString("z")
erro = request.QueryString("e")
vindo = request.QueryString("vd")
obr = request.QueryString("o")

if vindo="crmt" then
dados= split(obr, "_" )
unidade= dados(0)
curso= dados(1)
co_etapa= dados(2)
turma= dados(3)
end if
obr=cod
		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1
		
		Set CON2 = Server.CreateObject("ADODB.Connection") 
		ABRIR2 = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON2.Open ABRIR2
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0		

 call navegacao (CON,chave,nivel)
navega=Session("caminho")	

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



		Set RS_turma = Server.CreateObject("ADODB.Recordset")
		SQL_turma = "SELECT * FROM TB_Matriculas WHERE NU_Ano="& ano_letivo &" AND NU_Unidade ="& unidade &" AND CO_Curso='"&curso&"' AND CO_Etapa='"&etapa&"' AND CO_Turma='"&turma&"'"
		RS_turma.Open SQL_turma, CON1

		conta_alunos=0
		while not RS_turma.eof
		codigo_turma = RS_turma("CO_Matricula")
		
			if conta_alunos=0 then
			alunos_turma=codigo_turma
			else
			alunos_turma=codigo_turma&","&alunos_turma
			end if
		
		conta_alunos=conta_alunos+1
		RS_turma.MOVENEXT
		wend





ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

if mes<10 then
meswrt="0"&mes
else
meswrt=mes
end if
if min<10 then
minwrt="0"&min
else
minwrt=min
end if

data_proc = dia &"/"& meswrt &"/"& ano
horario = hora & ":"& minwrt


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

function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function checksubmit()
{
  if (document.busca.busca1.value != "" && document.busca.busca2.value != "")
  {    alert("Por favor digite SOMENTE uma opção de busca!")
    document.busca.busca1.focus()
    return false
  }
    if (document.busca.busca1.value == "" && document.busca.busca2.value == "")
  {    alert("Por favor digite uma opção de busca!")
    document.busca.busca1.focus()
    return false
  }
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
//-->
</script></head>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0"> 
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr> 
            
    <td height="10" colspan="5" class="tb_caminho"><font class="style-caminho"> 
      <%
	  response.Write(navega)

%>
       
    </td>
          </tr>
            <tr> 
              
    <td height="10" colspan="5" valign="top"> 
      <%call mensagens(nivel,636,0,0) %>
    </td>
			  </tr>			  
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
                      Matr&iacute;cula: </div></td>
                  <td width="9%" height="10"><font class="form_corpo"> 
                    <input name="cod" type="hidden" value="<%=codigo%>">
                    <%response.Write(codigo)%>
                    </td>
                  <td width="6%" height="10"> <div align="right"><font class="form_dado_texto"> 
                      Nome: </div></td>
                  <td width="66%" height="10"><font class="form_corpo"> 
                    <%response.Write(nome_prof)%>
                    <input name="nome" type="hidden" class="textInput" id="nome"  value="<%response.Write(nome_prof)%>" size="75" maxlength="50">
                    &nbsp;</td>
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
                  <td width="33" height="10"> <div align="center"> 
                      <%
call GeraNomes("PORT",unidade,curso,etapa,CON0)
no_unidades = session("no_unidades")
no_grau = session("no_grau")
no_serie = session("no_serie")
%>
                      Ano</div></td>
                  <td width="81" height="10"> <div align="center">Matr&iacute;cula</div></td>
                  <td width="75" height="10"> <div align="center">Cancelamento</div></td>
                  <td width="86" height="10"> <div align="center"> Situa&ccedil;&atilde;o</div></td>
                  <td width="113" height="10"> <div align="center">Unidade</div></td>
                  <td width="133" height="10"> <div align="center">Curso</div></td>
                  <td width="85" height="10"> <div align="center"> Etapa</div></td>
                  <td width="90" height="10"> <div align="center">Turma </div></td>
                  <td width="54" height="10"> <div align="center">Chamada</div></td>
                </tr>
                <tr class="tb_corpo"
> 
                  <td width="33" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(ano_aluno)%>
                      </div></td>
                  <td width="81" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(rematricula)%>
                      </div></td>
                  <td width="75" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(encerramento)%>
                      </div></td>
                  <td width="86" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)%>
                      </div></td>
                  <td width="113" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_unidades)%>
                      </div></td>
                  <td width="133" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_grau)%>
                      </div></td>
                  <td width="85" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_serie)%>
                      </div></td>
                  <td width="90" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      </div></td>
                  <td width="54" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(cham)%>
                      </div></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td>&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td height="10" colspan="2" class="tb_tit"
>Avalia&ccedil;&otilde;es</td>
          </tr>
          <tr> 
            <td colspan="2"><table width="100%" border="0" align="center" cellspacing="0" class="tb_corpo"
>
                <tr> 
                  <td> 

                    <%		Set RS_tb = Server.CreateObject("ADODB.Recordset")
		SQL_tb = "SELECT * FROM TB_Da_Aula WHERE NU_Unidade="& unidade &" AND CO_Curso ='"& curso&"' AND CO_Etapa ='"& etapa&"' AND CO_Turma ='"& turma&"'"
		RS_tb.Open SQL_tb, CON2
if RS_tb.eof then
%>
                    <div align="center" class="form_dado_texto"> <%response.Write("<br><br><br><br><br>Não existe Boletim para este aluno!")%>
                    </div>
                    <%
else
notaFIL=RS_tb("TP_Nota")

	if notaFIL ="TB_NOTA_A" then
	CAMINHOn = CAMINHO_na
	
	elseif notaFIL="TB_NOTA_B" then
		CAMINHOn = CAMINHO_nb
	
	elseif notaFIL ="TB_NOTA_C" then
			CAMINHOn = CAMINHO_nc
	else
			response.Write("ERRO")
	end if
	

'if notaFIL="TB_NOTA_A" or notaFIL="TB_NOTA_B" or notaFIL="TB_NOTA_C" then			
	%>
                    
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr> 
                      <td width="320" rowspan="2" class="tb_subtit"> 
                        <div align="left"><strong>Disciplina</strong></div></td>
                      <td colspan="12" class="tb_tit"> <div align="center">Aproveitamento</div></td>
                      <td colspan="4" class="tb_tit"><div align="center">Freq&uuml;&ecirc;ncia 
                          (Faltas)</div></td>
                    </tr>
                    <tr> 
                      <td width="68" class="tb_subtit"> <div align="center">BIM 
                          1</div></td>
                      <td width="68" class="tb_subtit"> <div align="center">BIM 
                          2</div></td>
                      <td width="68" class="tb_subtit"><div align="center">M&eacute;dia 
                          Sem 1</div></td>
                      <td width="68" class="tb_subtit"><div align="center">Recup 
                          Sem</div></td>
                      <td width="68" class="tb_subtit"><div align="center">M&eacute;dia 
                          Sem 2</div></td>
                      <td width="68" class="tb_subtit"> <div align="center">BIM 
                          3</div></td>
                      <td width="68" class="tb_subtit"> <div align="center">BIM 
                          4</div></td>
                      <td width="68" class="tb_subtit"><div align="center">M&eacute;dia 
                          Sem 3</div></td>
                      <td width="68" class="tb_subtit"> <div align="center">M&eacute;dia 
                          Anual</div></td>
                      <td width="68" class="tb_subtit"> <div align="center">Recup 
                          Final </div></td>
                      <td width="68" class="tb_subtit"><div align="center">M&eacute;dia Final</div></td>
                      <td width="68" class="tb_subtit"><div align="center">Result</div></td>
                      <td width="68" class="tb_subtit"><div align="center">BIM 
                          1</div></td>
                      <td width="68" class="tb_subtit"> <div align="center">BIM 
                          2</div></td>
                      <td width="68" class="tb_subtit"> <div align="center">BIM 
                          3</div></td>
                      <td width="68" class="tb_subtit"> <div align="center">BIM 
                          4</div></td>
                    </tr>
                    <%
			rec_lancado="sim"
		
			Set RSprog = Server.CreateObject("ADODB.Recordset")
			SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim "
			RSprog.Open SQLprog, CON0
		
			check=2
				
			while not RSprog.EOF
			
					dividendo1=0
					divisor1=0
					dividendo2=0
					divisor2=0
					dividendo3=0
					divisor3=0
					dividendo4=0
					divisor4=0
					dividendo_ma=0
					divisor_ma=0
					dividendo5=0
					divisor5=0
					dividendo_mf=0
					divisor_mf=0
					dividendo6=0
					divisor6=0
					dividendo_rec=0
					divisor_rec=0
					res1="&nbsp;"
					res2="&nbsp;"
					res3="&nbsp;"
					m2="&nbsp;"
					m3="&nbsp;"
															
			verifica="ok"
			
				materia=RSprog("CO_Materia")
				mae=RSprog("IN_MAE")
				fil=RSprog("IN_FIL")
				in_co=RSprog("IN_CO")
				nu_peso=RSprog("NU_Peso")
				ordem=RSprog("NU_Ordem_Boletim")
				
						Set RS1a = Server.CreateObject("ADODB.Recordset")
						SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
						RS1a.Open SQL1a, CON0
							
						no_materia=RS1a("NO_Materia")
						
						if check mod 2 =0 then
						cor = "tb_fundo_linha_par" 
						cor2 = "tb_fundo_linha_impar" 				
						else 
						cor ="tb_fundo_linha_impar"
						cor2 = "tb_fundo_linha_par" 
						end if										

				'response.Write(materia&" "&mae&" "&in_co&"<br>")
				
				if mae=TRUE and fil=FALSE and in_co=FALSE and isnull(nu_peso) then
				f1_ac=0
				f2_ac=0
				f3_ac=0
				f4_ac=0
				mb1_ac=0
				mb2_ac=0
				mb3_ac=0
				mb4_ac=0
				divisor_mb1_ac=0
				divisor_mb2_ac=0
				divisor_mb3_ac=0
				divisor_mb4_ac=0								
				ms1_ac=0
				ms2_ac=0
				ms3_ac=0
				ms35_ac=0
				ms36_ac=0								
				ma_ac=0
				peso_ac=0
				ordem2=ordem+1
				tentativas=0
								
				
			
						Set RS1a = Server.CreateObject("ADODB.Recordset")
						SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia&"'"
						RS1a.Open SQL1a, CON0
							
						no_materia=RS1a("NO_Materia")
					
						if check mod 2 =0 then
						cor = "tb_fundo_linha_par" 
						cor2 = "tb_fundo_linha_impar" 				
						else 
						cor ="tb_fundo_linha_impar"
						cor2 = "tb_fundo_linha_par" 
						end if
					
							
							
						Set CON_N = Server.CreateObject("ADODB.Connection") 
						ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
						CON_N.Open ABRIRn
					

						for periodofil=1 to 5
											
								
							Set RSnFIL = Server.CreateObject("ADODB.Recordset")
							Set RS3 = Server.CreateObject("ADODB.Recordset")
							SQL_N = "Select * from "& notaFIL &" WHERE CO_Matricula = "& cod &" AND CO_Materia = '"& materia &"' AND NU_Periodo="&periodofil
							Set RS3 = CON_N.Execute(SQL_N)
						
						
						
							if RS3.EOF then
								if periodofil=1 then
								f1="&nbsp;"
								va_m31="&nbsp;"
								elseif periodofil=2 then
								f2="&nbsp;"
								va_m32="&nbsp;"
								va_rec_sem="&nbsp;"
								elseif periodofil=3 then
								f3="&nbsp;"
								va_m33="&nbsp;"
								elseif periodofil=4 then
								f4="&nbsp;"
								va_m34="&nbsp;"
								elseif periodofil=5 then
								va_m35="&nbsp;"
								elseif periodofil=6 then
								va_m36="&nbsp;"
								end if	
							else
								if periodofil=1 then
								f1=RS3("NU_Faltas")
								va_m31=RS3("VA_Media3")
								elseif periodofil=2 then
								f2=RS3("NU_Faltas")
								va_m32=RS3("VA_Media3")
								va_rec_sem=RS3("VA_Rec")
								elseif periodofil=3 then
								f3=RS3("NU_Faltas")
								va_m33=RS3("VA_Media3")
								elseif periodofil=4 then
								f4=RS3("NU_Faltas")
								va_m34=RS3("VA_Media3")
								elseif periodofil=5 then
								va_m35=RS3("VA_Media3")
								elseif periodofil=6 then
								va_m36=RS3("VA_Media3")
								end if
							end if
						NEXT
							
							if isnull(va_m31) or va_m31="&nbsp;"  or va_m31="" then
							dividendo1=0
							divisor1=0
							else
							dividendo1=va_m31
							divisor1=1
							end if	
							
							if isnull(va_m32) or va_m32="&nbsp;" or va_m32="" then
							dividendo2=0
							divisor2=0
							else
							dividendo2=va_m32
							divisor2=1
							end if
							
							if isnull(va_rec_sem) or va_rec_sem="&nbsp;"  or va_rec_sem="" then
							dividendorec=0
							divisorrec=0
							else
							dividendorec=va_rec_sem
							divisorrec=1
							end if
							
							if isnull(va_m33) or va_m33="&nbsp;" or va_m33="" then
							dividendo3=0
							divisor3=0
							else
							dividendo3=va_m33
							divisor3=1
							end if
							
							if isnull(va_m34) or va_m34="&nbsp;" or va_m34="" then
							dividendo4=0
							divisor4=0
							else
							dividendo4=va_m34
							divisor4=1
							end if														
							if isnull(va_m35) or va_m35="&nbsp;" or va_m35="" then
							nota_aux_m2_1="&nbsp;"
							'dividendo5=0
							'divisor5=0
							else
							nota_aux_m2_1=va_m35
							'dividendo5=va_m35
							'divisor5=1
							end if

							showapr1="s"
							showprova1="s"
							showapr2="s"
							showprova2="s"
							showapr3="s"
							showprova3="s"
							showapr4="s"
							showprova4="s"

							dividendo_ms1=dividendo1+dividendo2
							divisor_ms1=divisor1+divisor2
													
							if divisor_ms1<2 then
							ms1="&nbsp;"
							dividendoms1=0
							divisorms1=0
							else
							ms1=dividendo_ms1/divisor_ms1
								decimo = ms1 - Int(ms1)
								If decimo >= 0.5 Then
									nota_arredondada = Int(ms1) + 1
									ms1=nota_arredondada
								'elseIf decimo >= 0.25 Then
								'	nota_arredondada = Int(ms1) + 0.5
								'	ms1=nota_arredondada
								else
									nota_arredondada = Int(ms1)
									ms1=nota_arredondada											
								End If			
							ms1 = formatNumber(ms1,0)
							dividendoms1=ms1
							divisorms1=1
							end if
							

							
							if divisorrec=0 then
								ms2=ms1
								if ms2="&nbsp;" then
									dividendoms2=0
									divisorms2=0
									dividendo_anual_ms2=0
									divisor_anual_ms2=0
								else
									dividendoms2=ms2
									divisorms2=1						
									dividendo_anual_ms2=ms2
									divisor_anual_ms2=1
								end if
							else
								dividendo_ms2=dividendoms1+dividendorec
								divisor_ms2=divisorms1+divisorrec
																						
								ms2=dividendo_ms2/divisor_ms2
'response.Write(ms2&"+"&dividendoms1&"+"&divisor_ms2)
ms2=ms2*1	
ms1=ms1*1							
								if ms2<ms1 then
									ms2=ms1								
								end if
									decimo = ms2 - Int(ms2)
									If decimo >= 0.5 Then
										nota_arredondada = Int(ms2) + 1
										ms2=nota_arredondada
									'elseIf decimo >= 0.25 Then
									'	nota_arredondada = Int(ms2) + 0.5
									'	ms2=nota_arredondada
									else
										nota_arredondada = Int(ms2)
										ms2=nota_arredondada											
									End If
								ms2 = formatNumber(ms2,0)																
								dividendo_anual_ms2=ms2
								divisor_anual_ms2=1
							end if
							
							dividendo_ms3=dividendo3+dividendo4
							divisor_ms3=divisor3+divisor4
							
							if divisor_ms3<2 then
							ms3="&nbsp;"
							dividendo_anual_ms3=0
							divisor_anual_ms3=0					
							else
							ms3=dividendo_ms3/divisor_ms3
								decimo = ms3 - Int(ms3)
								If decimo >= 0.5 Then
									nota_arredondada = Int(ms3) + 1
									ms3=nota_arredondada
								'elseIf decimo >= 0.25 Then
								'	nota_arredondada = Int(ms3) + 0.5
								'	ms3=nota_arredondada
								else
									nota_arredondada = Int(ms3)
									ms3=nota_arredondada											
								End If
								ms3 = formatNumber(ms3,0)								
							dividendo_anual_ms3=ms3
							divisor_anual_ms3=1						
							end if					
								dividendo_anual_ms2=dividendo_anual_ms2*1
								dividendo_anual_ms3=dividendo_anual_ms3*1
								divisor_anual_ms2=divisor_anual_ms2*1
								divisor_anual_ms3=divisor_anual_ms3*1		
							dividendo_ma=dividendo_anual_ms2+dividendo_anual_ms3
							divisor_ma=divisor_anual_ms2+divisor_anual_ms3
							
							'response.Write(dividendo_ma&"<<")
							
							if divisor_ma<2 then
							ma="&nbsp;"
							else
							ma=dividendo_ma/divisor_ma
							end if
							
							if ma="&nbsp;" then
							else
							'mf=mf/10
								decimo = ma - Int(ma)
		
									If decimo >= 0.5 Then
										nota_arredondada = Int(ma) + 1
										ma=nota_arredondada
									'elseIf decimo >= 0.25 Then
									'	nota_arredondada = Int(ma) + 0.5
									'	ma=nota_arredondada								
									Else
										nota_arredondada = Int(ma)
										ma=nota_arredondada					
									End If
									
								ma = formatNumber(ma,0)
							end if
		
							if ma="&nbsp;" then
								media_final="&nbsp;"
								resultado_final="&nbsp;"							
							else	
							
								if nota_aux_m2_1="&nbsp;" then
									tipo_calculo="anual"
								else
									tipo_calculo="final"
								end if	
								if isnumeric(ma) then	
									ma=ma/10
								end if	
								
								if isnumeric(nota_aux_m2_1) then	
									nota_aux_m2_1=nota_aux_m2_1/10
								end if																	

								resultado=regra_aprovacao(curso,etapa,ma,nota_aux_m2_1,"&nbsp;","&nbsp;","&nbsp;",tipo_calculo)							
								resultado_aluno = split(resultado,"#!#")
								media_final=resultado_aluno(0)
								if isnumeric(media_final) then
									media_final=media_final*10
								end if
								resultado_final=resultado_aluno(1)								
							end if
							
							
'response.Write(Session("resultado_1") &" - "& Session("resultado_2") &" - "& Session("resultado_3")&"<<")
			%>
                    <tr> 
                      <td width="320" class="<%response.Write(cor)%>"> 
                        <%response.Write(no_materia)
						  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
						  %>
                      </td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%
							if showapr1="s" then																	
							response.Write(va_m31)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%
							if showapr1="s" then												
							response.Write(va_m32)						
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
                          <%
							if showapr3="s" then					
							response.Write(ms1)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
                          <%
							if showapr1="s" then					
							response.Write(va_rec_sem)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
                          <%
							if showapr3="s" then					
							response.Write(ms2)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%
							if showapr1="s" then					
							response.Write(va_m33)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%
							if showapr1="s" then					
							response.Write(va_m34)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%
							if showapr3="s" then					
							response.Write(ms3)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%
							if showapr3="s" then
							response.Write(ma)
							else
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%
							if showapr2="s" and showprova2="s" then												
							response.Write(va_m35)
							end if
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center"><%				
								response.Write(media_final)
							%></div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                          <%				
								response.Write(resultado_final)
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                          <%
							response.Write(f1)
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                          <%
							response.Write(f2)
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                          <%
							response.Write(f3)
							%>
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                          <%
							response.Write(f4)
							%>
                        </div></td>
                    </tr>
                    <%
					check=check+1
					else%>
                    <tr> 
                      <td width="320" class="<%response.Write(cor)%>"><%response.Write(no_materia)
						  'response.Write("("&unidade&"-"&curso&"-"&etapa&"-"&turma&"-"&ma&"-"&nota_aux1&"-"&mf&"-"&nota_aux2&"-"&rec&")")
						  %></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"><div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor)%>">&nbsp;</td>
                      <td width="68" class="<%response.Write(cor)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                        </div></td>
                      <td width="68" class="<%response.Write(cor2)%>"> <div align="center"> 
                        </div></td>
                    </tr>
                    <%
					check=check+1
					end if	
			RSprog.MOVENEXT
			wend
			
			'Set RSF = Server.CreateObject("ADODB.Recordset")
			'SQL_N = "Select * from TB_Frequencia_Periodo WHERE CO_Matricula = "& cod
			'Set RSF = CON_N.Execute(SQL_N)
			
			'if RSF.eof THEN
			'f1="&nbsp;"
			'f2="&nbsp;"
			'f3="&nbsp;"
			'f4="&nbsp;"			
			'else	
			'f1=RSF("NU_Faltas_P1")
			'f2=RSF("NU_Faltas_P2")
			'f3=RSF("NU_Faltas_P3")
			'f4=RSF("NU_Faltas_P4")		
			'END IF		
			%>
                    <tr valign="bottom"> 
                      <td height="20"> <div align="right"> </div></td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td>&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                      <td height="20">&nbsp;</td>
                    </tr>
                  </table>
<%
end if					
					%>
                  </td>
                </tr>
              </table></td>
          </tr>
        </table></td>
    </tr>
  <tr>
    <td height="40" colspan="5" valign="top"><img src="../../../../img/rodape.jpg" width="1000" height="40"></td>
  </tr>
</table>

</body>
<%call GravaLog (chave,obr)%>
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