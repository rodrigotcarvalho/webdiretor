<%On Error Resume Next%>
<!--#include file="../../inc/connect_wf.asp"-->
<!--#include file="../../inc/connect_al.asp"-->
<!--#include file="../../inc/connect_pf.asp"-->
<!--#include file="../../inc/connect_p.asp"-->
<!--#include file="../../inc/funcoes.asp"-->
<!--#include file="../../inc/funcoes2.asp"-->

<%
nivel=2
tp=session("tp")
nome = session("nome") 
co_user = session("co_user")
opt=request.QueryString("opt")

if opt="1" then
'periodo_check=request.form("periodo")
cod= Session("aluno_selecionado")
else
cod= Session("aluno_selecionado")
'periodo_check=1
end if
cod= Session("aluno_selecionado")

obr=cod&"?"&periodo_check

 	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0

 	Set CON = Server.CreateObject("ADODB.Connection") 
	ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON.Open ABRIR
	
	Set CON1 = Server.CreateObject("ADODB.Connection") 
	ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON1.Open ABRIR1

		Set CON4 = Server.CreateObject("ADODB.Connection") 
		ABRIR4 = "DBQ="& CAMINHO_pf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON4.Open ABRIR4

	SQL2 = "select * from TB_Usuario where CO_Usuario = " & cod 
	set RS2 = CON.Execute (SQL2)
	
nome_aluno= RS2("NO_Usuario")

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



%>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
<title>Web Família</title>
<link href="../../estilo.css" rel="stylesheet" type="text/css" />
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
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>

<body onload="MM_preloadImages(<%response.Write(swapload)%>)">
<form action="index.asp?opt=1" method="post">
  <table width="1000" height="1039" border="0" align="center" cellpadding="0" cellspacing="0" class="tabela_principal">
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
 mes = "março"
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
data_compara=data
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
            <td width="172" height="144" rowspan="4" valign="top" class="tabela_menu"><p>&nbsp;</p>
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
            <td height="5" colspan="2"><p><img src="../../img/linha-pontilhada_grande.gif" alt="" width="828" height="5" /></p></td>
          </tr>
		       <tr class="tabela_menu">
        <td height="19" colspan="2">&nbsp;</td>
      </tr> 
          <tr class="tabela_menu"> 
            <td height="832" colspan="2" valign="top"> <div align="left"><img src="../../img/extrato.jpg" width="700" height="30"> 
              </div>
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
                        <td width="19%" height="10"> <div align="right"><font class="style3"> 
                            Matr&iacute;cula: </font></div></td>
                        <td width="9%" height="10"><font class="style1"> 
                          <input name="cod" type="hidden" value="<%=cod%>">
                          <%response.Write(cod)%>
                          </font></td>
                        <td width="6%" height="10"> <div align="right"><font class="style3"> 
                            Nome: </font></div></td>
                        <td width="66%" height="10"><font class="style1"> 
                          <%response.Write(nome_aluno)%>
                          <input name="nome2" type="hidden" class="textInput" id="nome2"  value="<%response.Write(nome_aluno)%>" size="75" maxlength="50">
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
                      <tr class="style3"> 
                        <td width="34" height="10"> <div align="center"> 
                            <%					  
call GeraNomes("PORT",unidade,curso,etapa,CON0)
no_unidade = session("no_unidade")
no_curso = session("no_curso")
no_etapa= session("no_etapa")
%>
                            ANO</div></td>
                        <td width="74" height="10"> <div align="center">MATR&Iacute;CULA</div></td>
                        <td width="96" height="10"> <div align="center">CANCELAMENTO</div></td>
                        <td width="83" height="10"> <div align="center"> SITUA&Ccedil;&Atilde;O</div></td>
                        <td width="81" height="10"> <div align="center">UNIDADE</div></td>
                        <td width="111" height="10"> 
                          <div align="center">CURSO</div></td>
                        <td width="63" height="10"> <div align="center"> ETAPA</div></td>
                        <td width="66" height="10"> <div align="center">TURMA</div></td>
                        <td width="60" height="10"> 
                          <div align="center">CHAMADA</div></td>
                      </tr>
                      <tr class="tb_corpo"
> 
                        <td width="34" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(ano_aluno)%>
                            </font></div></td>
                        <td width="74" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(rematricula)%>
                            </font></div></td>
                        <td width="96" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(encerramento)%>
                            </font></div></td>
                        <td width="83" height="10"> <div align="center"> <font class="style1"> 
                            <%
					
		Set RSCONTST = Server.CreateObject("ADODB.Recordset")
		SQLCONTST = "SELECT * FROM TB_Situacao_Aluno WHERE CO_Situacao='"&situacao&"'"
		RSCONTST.Open SQLCONTST, CON0
							
				no_situacao = RSCONTST("TX_Descricao_Situacao")	
					response.Write(no_situacao)%>
                            </font></div></td>
                        <td width="81" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(no_unidade)%>
                            </font></div></td>
                        <td width="111" height="10"> 
                          <div align="center"> <font class="style1"> 
                            <%response.Write(no_curso)%>
                            </font></div></td>
                        <td width="63" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(no_etapa)%>
                            </font></div></td>
                        <td width="66" height="10"> <div align="center"> <font class="style1"> 
                            <%response.Write(turma)%>
                            </font></div></td>
                        <td width="60" height="10"> 
                          <div align="center"> <font class="style1"> 
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
                      <tr class="style3"> 
                        <td width="138"> <div align="center">DATA VENCIMENTO</div></td>
                        <td width="138"> <div align="center">SERVI&Ccedil;O</div></td>
                        <td width="138"> <div align="center">VALOR A PAGAR</div></td>
                        <td width="138"> <div align="center">VALOR PAGO</div></td>
                        <td width="138"> <div align="center">DATA PAGAMENTO</div></td>
                        <td width="138"> <div align="center">SITUA&Ccedil;&Atilde;O</div></td>
                      </tr>
                      <%		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4= "SELECT * FROM TB_Posicao WHERE CO_Matricula_Escola ="& cod
		RS4.Open SQL4, CON4
		
if RS4.EOF THEN
%>
                      <tr> 
                        <td colspan="6"> <div align="center"><font class="style1"> 
                            Não existem lançamentos financeiros para este aluno.</font></div></td>
                      </tr>
                      <%else
check = 2
while not RS4.EOF
compromisso=RS4("VA_Compromisso")
da_vencimento=RS4("DA_Vencimento")
realizado=RS4("VA_Realizado") 
da_realizado=RS4("DA_Realizado")
nome_lanc=RS4("NO_Lancamento")

if realizado = 0 then
realizado=""
else
realizado=FormatCurrency(realizado)
end if

venc_split=split(da_vencimento,"/")
dia_venc=venc_split(0)
mes_venc=venc_split(1)
ano_venc=venc_split(2)
venc=mes_venc&"/"&dia_venc&"/"&ano_venc

if dia_venc<10 then
dia_venc="0"&dia_venc
else
dia_venc=dia_venc
end if

if mes_venc<10 then
mes_venc="0"&mes_venc
else
mes_venc=mes_venc
end if

da_vencimento_show=dia_venc&"/"&mes_venc&"/"&ano_venc

venc=replace(da_vencimento,"/","$wxg$adn$")

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

if dia_real<10 then
dia_real="0"&dia_real
else
dia_real=dia_real
end if

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
 -->                                <% response.Write(da_vencimento_show)%>
<!--                             </a> 
  -->                             <%elseif situacao="em atraso" then %>
<!--                            <a href="#" class="menu_lista" onClick="MM_openBrWindow('../segvia/bloqueto.asp?c=<%=cod%>&amp;m=<%=mes_venc%>&amp;v=<%=venc%>&amp;opt=c','','width=700,height=100')"> 
 -->                             <% response.Write(da_vencimento_show)%>
 <!--                           </a> 
  -->                               <%else
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
%>                      <tr class="<% = cor %>"> 
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                        <td>&nbsp;</td>
                      </tr>
                      <tr class="<% = cor %>"> 
                        <td colspan="6"><table width="200" height="20" border="0" align="center" cellpadding="0" cellspacing="0">
                            <tr> 
                              <td class="tb_tit"><div align="center"><a href="#" class="impressao" onClick="MM_openBrWindow('imprime.asp?obr=<%=obr%>','','menubar=yes,width=1000,height=450')">Versão 
                                  para impressão</a></div></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table></td>
                </tr>
              </table></td>
          </tr>
        </table></td>
  </tr>
  <tr>
      <td width="1000" height="41"><img src="../../img/rodape.jpg" width="1000" height="41" /></td>
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
response.redirect("../../inc/erro.asp")
end if
%>