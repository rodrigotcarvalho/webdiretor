<%' Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->


<%
opt = request.QueryString("opt")
cod= request.QueryString("cod_cons")
ano_letivo = session("ano_letivo")
co_usr = session("co_user")
nivel=4
obr=cod&"?"&ano_letivo
Session("dia_de")=""
Session("dia_de")=""
Session("dia_ate")=""
Session("mes_ate")=""
Session("unidade")=""
Session("curso")=""
Session("etapa")=""
Session("turma")=""


	nvg=session("chave")
	session("chave")=nvg
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
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
    	Set CON_WF = Server.CreateObject("ADODB.Connection") 
		ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_WF.Open ABRIR_WF	

	
		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

	
	
		'response.Write "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
	
		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4		

 call navegacao (CON,nvg,nivel)
navega=Session("caminho")	

 Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Autoriz_Usuario_Grupo Where CO_Usuario = "&co_usr
		RS2.Open SQL2, CON
		
if RS2.EOF then

else		
co_grupo=RS2("CO_Grupo")
End if



	SQL2 = "select * from TB_Alunos where CO_Matricula = " & cod 
	set RS2 = CON1.Execute (SQL2)
	
nome_aluno= RS2("NO_Aluno")

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

data = dia &"/"& meswrt &"/"& ano
data_compara=data
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
<script type="text/javascript" src="../../../../js/global.js"></script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_reloadPage(init) {  //reloads the window if Nav4 resized
  if (init==true) with (navigator) {if ((appName=="Netscape")&&(parseInt(appVersion)==4)) {
    document.MM_pgW=innerWidth; document.MM_pgH=innerHeight; onresize=MM_reloadPage; }}
  else if (innerWidth!=document.MM_pgW || innerHeight!=document.MM_pgH) location.reload();
}
MM_reloadPage(true);function submitforminterno()  
{
   var f=document.forms[0]; f.submit(); 
	  
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
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
  <link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif')">
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
      <tr>                   
    <td height="10"> 
      <%	call mensagens(nivel,636,0,0) 
%>
</td></tr>
<tr>

            <td valign="top"> 
<%
mes = DatePart("m", now) 
dia = DatePart("d", now) 



dia=dia*1
mes=mes*1
%>				
<FORM METHOD="POST" ENCTYPE="multipart/form-data" ACTION="../apf/upload.asp?opt=f&al=<%=ano_letivo%>" target="_parent">
        <table width="100%" border="0" align="center" cellspacing="0" class="tb_corpo"
>
          <tr> 
            <td width="684" class="tb_tit"
>Dados Escolares</td>
            <td width="140" class="tb_tit"
> </td>
          </tr>
          <tr> 
            <td height="10"> <table width="100%" border="0" cellspacing="0">
                <tr> 
                  <td width="19%" height="10"> <div align="right"><font class="form_dado_texto"> 
                      <strong>Matr&iacute;cula:</strong> </font></div></td>
                  <td width="9%" height="10" ><font class="form_dado_texto"> 
                    <input name="cod" type="hidden" value="<%=cod%>">
                    <%response.Write(cod)%>
                    </font></td>
                  <td width="6%" height="10"> <div align="right"><font class="form_dado_texto"> 
                      <strong>Nome: </strong></font></div></td>
                  <td width="66%" height="10"><font class="form_dado_texto"> 
                    <%response.Write(nome_aluno)%>
                    <input name="nome" type="hidden" class="textInput" id="nome"  value="<%response.Write(nome_aluno)%>" size="75" maxlength="50">
                    &nbsp;</font></td>
                </tr>
              </table></td>
            <td rowspan="2" valign="top"><div align="center"> </div></td>
          </tr>
          <tr> 
            <td height="10" bgcolor="#FFFFFF">&nbsp;</td>
          </tr>
          <tr> 
            <td colspan="2"><table width="100%" border="0" cellspacing="0">
                <tr class="tb_subtit"> 
                  <td width="34" height="10"> <div align="center"> 
                      <%					  
	no_unidade= GeraNomes("U",unidade,variavel2,variavel3,variavel4,variavel5,CON0,outro)
	no_curso=GeraNomes("C",curso,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
	no_etapa=GeraNomes("E",etapa,variavel2,variavel3,variavel4,variavel5,CON0,outro) 	
%>
                      ANO</div></td>
                  <td width="74" height="10"> <div align="center">MATR&Iacute;CULA</div></td>
                  <td width="96" height="10"> <div align="center">CANCELAMENTO</div></td>
                  <td width="83" height="10"> <div align="center"> SITUA&Ccedil;&Atilde;O</div></td>
                  <td width="81" height="10"> <div align="center">UNIDADE</div></td>
                  <td width="111" height="10"> <div align="center">CURSO</div></td>
                  <td width="63" height="10"> <div align="center"> ETAPA</div></td>
                  <td width="66" height="10"> <div align="center">TURMA</div></td>
                  <td width="60" height="10"> <div align="center">CHAMADA</div></td>
                </tr>
                <tr class="tb_corpo"
> 
                  <td width="34" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(ano_aluno)%>
                      </font></div></td>
                  <td width="74" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(rematricula)%>
                      </font></div></td>
                  <td width="96" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(encerramento)%>
                      </font></div></td>
                  <td width="83" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)%>
                      </font></div></td>
                  <td width="81" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_unidade)%>
                      </font></div></td>
                  <td width="111" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_curso)%>
                      </font></div></td>
                  <td width="63" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(no_etapa)%>
                      </font></div></td>
                  <td width="66" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(turma)%>
                      </font></div></td>
                  <td width="60" height="10"> <div align="center"> <font class="form_dado_texto"> 
                      <%response.Write(cham)%>
                      </font></div></td>
                </tr>
              </table></td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr> 
            <td height="10" colspan="2" class="tb_tit"
>Extrato</td>
          </tr>
          <tr> 
            <td colspan="2"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr class="tb_subtit"> 
                  <td width="138"> 
                    <div align="center">DATA VENCIMENTO</div></td>
                  <td width="138"> 
                    <div align="center">SERVI&Ccedil;O</div></td>
                  <td width="138"> 
                    <div align="center">VALOR A PAGAR</div></td>
                  <td width="138"> 
                    <div align="center">VALOR PAGO</div></td>
                  <td width="138"> 
                    <div align="center">DATA PAGAMENTO</div></td>
                  <td width="138"> 
                    <div align="center">SITUA&Ccedil;&Atilde;O</div></td>
                </tr>
                <%		
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4= "SELECT * FROM TB_Posicao WHERE CO_Matricula_Escola ="& cod
		RS4.Open SQL4, CON4
		
if RS4.EOF THEN
%>
                <tr> 
                  <td colspan="6"> <div align="center"><font class="form_dado_texto"> <br>
                      <br>
                      <br>
                      <br>
                      <br>
                      Não existem lançamentos financeiros para este aluno.<br>
                      <br>
                      <br>
                      <br>
                      <br>
                      </font></div></td>
                </tr>
                <%else
check = 2
while not RS4.EOF
compromisso=RS4("VA_Compromisso")
da_vencimento=RS4("DA_Vencimento")
realizado=RS4("VA_Realizado") 
da_realizado=RS4("DA_Realizado")
nome_lanc=RS4("NO_Lancamento")

if realizado = 0 or isnull(realizado) then
realizado=""
else
realizado=FormatCurrency(realizado)
end if

venc_split=split(da_vencimento,"/")
dia_venc=venc_split(0)
mes_venc=venc_split(1)
ano_venc=venc_split(2)
venc=mes_venc&"/"&dia_venc&"/"&ano_venc

dia_venc = dia_venc*1
if dia_venc<10 then
dia_venc="0"&dia_venc
else
dia_venc=dia_venc
end if

mes_venc = mes_venc*1
if mes_venc<10 then
mes_venc="0"&mes_venc
else
mes_venc=mes_venc
end if

da_vencimento_show=dia_venc&"/"&mes_venc&"/"&ano_venc

venc=replace(da_vencimento,"/","$wxg$adn$")
'RESPONSE.Write(data_compara&"<<")
if isnull(da_realizado) then
	d_diff=DateDiff("d",data_compara,da_vencimento)
	situacao="aberto"
	da_realizado_show=""
else
	real_split=split(da_realizado,"/")
	dia_real=real_split(0)
	mes_real=real_split(1)
	ano_real=real_split(2)
	real=mes_real&"/"&dia_real&"/"&ano_real
	
	dia_real = dia_real*1
	if dia_real<10 then
		dia_real="0"&dia_real
	else
		dia_real=dia_real
	end if
	
	mes_real=mes_real*1
	if mes_real<10 then
		mes_real="0"&mes_real
	else
		mes_real=mes_real
	end if
	
	da_realizado_show=dia_real&"/"&mes_real&"/"&ano_real
	
	d_diff=DateDiff("d",da_realizado,da_vencimento)
	situacao="pago"
end if
'response.Write(da_vencimento&" - "&da_realizado&" = "&d_diff&"<BR>")
if isnull(da_realizado) and d_diff<0 then
  cor = "tb_fundo_linha_atraso" 
  situacao="em atraso"
else
 if check mod 2 =0 then
  cor = "tb_fundo_linha_ok" 
  else cor ="tb_fundo_linha_ok"
  end if  
end if
%>
                <tr class="<% = cor %>"> 
                  <td width="138"> <div align="center"> 
                      <% if situacao="aberto" then %>
                      <!--                            <a href="#" class="menu_sublista" onClick="MM_openBrWindow('../segvia/bloqueto.asp?c=<%=cod%>&amp;m=<%=mes_venc%>&amp;v=<%=venc%>&amp;opt=c','','width=700,height=100')"> 
 -->
                      <% response.Write(da_vencimento_show)%>
                      <!--                             </a> 
  -->
                      <%elseif situacao="em atraso" then %>
                      <!--                            <a href="#" class="menu_lista" onClick="MM_openBrWindow('../segvia/bloqueto.asp?c=<%=cod%>&amp;m=<%=mes_venc%>&amp;v=<%=venc%>&amp;opt=c','','width=700,height=100')"> 
 -->
                      <% response.Write(da_vencimento_show)%>
                      <!--                           </a> 
  -->
                      <%else
 response.Write(da_vencimento_show)
end if%>
                    </div></td>
                  <td width="138"> <div align="center"> 
                      <% response.Write(nome_lanc)%>
                    </div></td>
                  <td width="138"> <div align="center"> 
                      <% response.Write(FormatCurrency(compromisso))%>
                    </div></td>
                  <td width="138"> <div align="center"> 
                      <% response.Write(realizado)%>
                    </div></td>
                  <td width="138"> <div align="center"> 
                      <% response.Write(da_realizado_show)%>
                    </div></td>
                  <td width="138"> <div align="center"> 
                      <% response.Write(situacao)%>
                    </div></td>
                </tr>
                <%RS4.MOVENEXT
WEND
END IF
%>
                <tr class="<% = cor %>"> 
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                  <td>&nbsp;</td>
                </tr>
                <tr class="<% = cor %>"> 
                  <td colspan="6">&nbsp;</td>
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
response.redirect("../../../../inc/erro.asp")
end if
%>