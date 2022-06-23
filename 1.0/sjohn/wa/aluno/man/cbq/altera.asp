<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->
<!--#include file="../../../../inc/funcoes6.asp"-->
<!--#include file="../../../../inc/bd_grade.asp"-->
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

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
 call navegacao (CON,chave,nivel)
navega=Session("caminho")		

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

obr=cod&"$!$"&unidade&"$!$"&curso&"$!$"&etapa&"$!$"&turma

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
</script>
</head>
<body link="#CC9900" vlink="#CC9900" alink="#CC9900" background="../../../../img/fundo.gif" leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%call cabecalho(nivel)
%>
<table width="1000" height="650" border="0" align="center" cellspacing="0" background="../../../../img/fundo_interno.gif" bgcolor="#FFFFFF">
  <tr>
    <td height="10" colspan="5" class="tb_caminho"><font class="style-caminho">
      <%
	  response.Write(navega)

%></td>
  </tr>
  <tr>
    <td height="10" colspan="5" valign="top"><%call mensagens(nivel,665,0,0) %></td>
  </tr>
  <tr>
    <td valign="top"><table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo"
>
        <tr>
          <td width="653" class="tb_tit"
>Dados Escolares</td>
          <td width="113" class="tb_tit"
></td>
        </tr>
        <tr>
          <td height="10"><table width="100%" border="0" cellspacing="0">
              <tr>
                <td width="19%" height="10"><div align="right"><font class="form_dado_texto"> Matr&iacute;cula: </div></td>
                <td width="9%" height="10"><font class="form_corpo">
                  <input name="cod" type="hidden" value="<%=codigo%>">
                  <%response.Write(codigo)%></td>
                <td width="6%" height="10"><div align="right"><font class="form_dado_texto"> Nome: </div></td>
                <td width="66%" height="10"><font class="form_corpo">
                  <%response.Write(nome_prof)%>
                  <input name="nome" type="hidden" class="textInput" id="nome"  value="<%response.Write(nome_prof)%>" size="75" maxlength="50">
                  &nbsp;</td>
              </tr>
            </table></td>
          <td valign="top">&nbsp;</td>
        </tr>
        <tr>
          <td height="10" bgcolor="#FFFFFF">&nbsp;</td>
          <td valign="top" bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
        <tr>
          <td colspan="2"><table width="100%" border="0" cellspacing="0">
              <tr class="tb_subtit">
                <td width="33" height="10"><div align="center">
                    <%
call GeraNomes("PORT",unidade,curso,etapa,CON0)
no_unidades = session("no_unidades")
no_grau = session("no_grau")
no_serie = session("no_serie")
%>
                    Ano</div></td>
                <td width="81" height="10"><div align="center">Matr&iacute;cula</div></td>
                <td width="75" height="10"><div align="center">Cancelamento</div></td>
                <td width="86" height="10"><div align="center"> Situa&ccedil;&atilde;o</div></td>
                <td width="113" height="10"><div align="center">Unidade</div></td>
                <td width="133" height="10"><div align="center">Curso</div></td>
                <td width="85" height="10"><div align="center"> Etapa</div></td>
                <td width="90" height="10"><div align="center">Turma </div></td>
                <td width="54" height="10"><div align="center">Chamada</div></td>
              </tr>
              <tr class="tb_corpo"
>
                <td width="33" height="10"><div align="center"> <font class="form_dado_texto">
                    <%response.Write(ano_aluno)%>
                  </div></td>
                <td width="81" height="10"><div align="center"> <font class="form_dado_texto">
                    <%response.Write(rematricula)%>
                  </div></td>
                <td width="75" height="10"><div align="center"> <font class="form_dado_texto">
                    <%response.Write(encerramento)%>
                  </div></td>
                <td width="86" height="10"><div align="center"> <font class="form_dado_texto">
                    <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)%>
                  </div></td>
                <td width="113" height="10"><div align="center"> <font class="form_dado_texto">
                    <%response.Write(no_unidades)%>
                  </div></td>
                <td width="133" height="10"><div align="center"> <font class="form_dado_texto">
                    <%response.Write(no_grau)%>
                  </div></td>
                <td width="85" height="10"><div align="center"> <font class="form_dado_texto">
                    <%response.Write(no_serie)%>
                  </div></td>
                <td width="90" height="10"><div align="center"> <font class="form_dado_texto">
                    <%response.Write(turma)%>
                  </div></td>
                <td width="54" height="10"><div align="center"> <font class="form_dado_texto">
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
>Avalia&ccedil;&otilde;es Qualitativas</td>
        </tr>
        <tr>
          <td colspan="2">
          <%
dados_periodo =  periodos(periodo, "num")
total_periodo = split(dados_periodo,"#!#") 
notas_a_lancar = ubound(total_periodo)-2	

	%>
                  <table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="252" class="tb_subtit" align="left">Disciplina</td>
    <%for i=0 to notas_a_lancar
		sigla_periodo =  periodos(total_periodo(i), "sigla")
	%>
    <td align="center" class="tb_subtit"><%response.Write(sigla_periodo)%></td>
    <%next%>
                    </tr>
                    <%
			rec_lancado="sim"
		
			Set RSprog = Server.CreateObject("ADODB.Recordset")
			SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' and IN_MAE= TRUE order by NU_Ordem_Boletim "
			RSprog.Open SQLprog, CON0
		
			check=2
			
			Set CON_N = Server.CreateObject("ADODB.Connection") 
			ABRIRn = "DBQ="& CAMINHO_nw & ";Driver={Microsoft Access Driver (*.mdb)}"
			CON_N.Open ABRIRn			
				
			while not RSprog.EOF
				co_materia=RSprog("CO_Materia")
				mae=RSprog("IN_MAE")
				fil=RSprog("IN_FIL")
				in_co=RSprog("IN_CO")
				nu_peso=RSprog("NU_Peso")
				ordem=RSprog("NU_Ordem_Boletim")
			
				Set RS1a = Server.CreateObject("ADODB.Recordset")
				SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&co_materia&"'"
				RS1a.Open SQL1a, CON0
					
				no_materia=RS1a("NO_Materia")
				co_materia_pr= RS1a("CO_Materia_Principal")
				
				if Isnull(co_materia_pr) then
					co_materia_pr= co_materia
				end if	
										
			
				if check mod 2 =0 then
				cor = "tb_fundo_linha_par" 
				cor2 = "tb_fundo_linha_impar" 				
				else 
				cor ="tb_fundo_linha_impar"
				cor2 = "tb_fundo_linha_par" 
				end if
				
     			%>
                <tr class="<%response.Write(cor)%>">
                      <td width="252"><%response.Write(no_materia)%></td>                
                <%				

			
				for j=0 to notas_a_lancar	
				    qtd_filhas=0
					acumula_valor = 0 
				    if mae=true then
							
						Set RS1a = Server.CreateObject("ADODB.Recordset")
						SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&co_materia_pr&"'"
						RS1a.Open SQL1a, CON0
							
						if RS1a.EOF then
							vetor_filhas = co_materia
						else
							while not RS1a.eof
								co_materia= RS1a("CO_Materia")
								if qtd_filhas=0 then
									vetor_filhas = co_materia
								else
									vetor_filhas = vetor_filhas&"#!#"&co_materia	
								end if
								qtd_filhas=qtd_filhas+1
							RS1a.MOVENEXT
							WEND				
						end if
					else
						vetor_filhas = co_materia						
					end if
					'response.Write(co_materia_pr&"-"&vetor_filhas&"<BR>")
					filhas = split(vetor_filhas,"#!#")
					wrk_calcula_medias = "S"
					for f=0 to ubound(filhas)
						co_materia = filhas(f)
					
								if j=0 then
									wrk_bd_nota_per = "VA_Ava1"
								elseif j=1 then
									wrk_bd_nota_per = "VA_Ava2"
								elseif j=2 then 			
									wrk_bd_nota_per = "VA_Ava3"
								elseif j=3 then 			 	
									wrk_bd_nota_per = "VA_Ava4"
								end if					
													
								Set RS3 = Server.CreateObject("ADODB.Recordset")
								SQL_N = "Select * from TB_Nota_W WHERE CO_Matricula = "& cod &" AND CO_Materia_Principal = '"& co_materia_pr &"' AND CO_Materia = '"& co_materia &"'"
								Set RS3 = CON_N.Execute(SQL_N)				
							
								if RS3.EOF then 
									valor="&nbsp;"
									wrk_calcula_medias = "N"									
								else						
									valor=RS3(wrk_bd_nota_per)	
									if isnull(valor) or valor="" then
										valor = "&nbsp;"
										wrk_calcula_medias = "N"																			
									else											
										if valor = "M" then
											valor = "MB"
										end if			   															
									end if
								end if	

							if qtd_filhas>0 then
								if valor = "I" then
									valor_convertido = 25
								elseif valor = "R" then
									valor_convertido = 50
								elseif valor = "B" then
									valor_convertido = 75
								else
									valor_convertido = 100
								end if		
								
								acumula_valor = acumula_valor+valor_convertido
							end if								
						next
						
						if acumula_valor>0 and qtd_filhas>0 and wrk_calcula_medias = "S" then
							valor = acumula_valor/qtd_filhas
							
							if valor <=25 then
								valor="I"
							elseif valor <=50 then
								valor="R"
							elseif valor <=75 then
								valor="B"
							else
								valor="MB"
							end if	
						else
							if wrk_calcula_medias = "N"  then								  
								valor="&nbsp;"
							end if	
						end if		
			%>
                    

                      <td align="center">
									<%response.Write(valor)%></td>
					<%NEXT%>
                    </tr>
                    <%
			check=check+1
			RSprog.MOVENEXT
			wend
			
%>                  </table>
</td>
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