<%On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/funcoes.asp"-->
<!--#include file="../../../../inc/funcoes2.asp"-->



<%
opt = request.QueryString("opt")

ano_letivo = Session("ano_letivo")
co_usr = session("co_user")
nivel=4

Session("dia_de")=""
Session("dia_de")=""
Session("dia_ate")=""
Session("mes_ate")=""
Session("unidade")=""
Session("curso")=""
Session("etapa")=""
Session("turma")=""


nvg = request.QueryString("nvg")
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


 call navegacao (CON,nvg,nivel)
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
function submitforminterno()  
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
						
						

						 function recuperarTurma(eTipo)
                                   {
// Criação do objeto XMLHTTP
                                               var oHTTPRequest = createXMLHTTP();
// Abrindo a solicitação HTTP. O primeiro parâmetro informa o método post/get
// O segundo parâmetro informa o arquivo solicitado que pode ser asp, php, txt, xml, etc.
// O terceiro parametro informa que a solicitacao nao assincrona,
// Para solicitação síncrona, o parâmetro deve ser false
                                               oHTTPRequest.open("post", "../../../../inc/executa.asp?opt=t", true);
// Para solicitações utilizando o método post, deve ser acrescentado este cabecalho HTTP
                                               oHTTPRequest.setRequestHeader("Content-Type", "application/x-www-form-urlencoded");
// A função abaixo é executada sempre que o estado do objeto muda (onreadystatechange)
                                               oHTTPRequest.onreadystatechange=function() {
// O valor 4 significa que o objeto já completou a solicitação
                                                           if (oHTTPRequest.readyState==4){
// Abaixo o texto é gerado no arquivo executa.asp e colocado no div
                                                                       var resultado_t= oHTTPRequest.responseText;
resultado_t = resultado_t.replace(/\+/g," ")
resultado_t = unescape(resultado_t)
document.all.divTurma.innerHTML = resultado_t																	   
                                                           }
                                               }
// Abaixo é enviada a solicitação. Note que a configuração
// do evento onreadystatechange deve ser feita antes do send.
                                               oHTTPRequest.send("e_pub=" + eTipo);
                                   }
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
function checksubmit()
{
  if (document.formulario.tipo.value == "nulo" )
  { alert("É necessário selecionar o tipo!")
	var combo = document.getElementById("tipo");
	combo.options[0].selected = "true";		
    return false
  } else if 
  (document.formulario.tipo.value == "it" && document.formulario.modalidade.value == "nulo")
  { alert("É necessário selecionar a modalidade!")
	var combo = document.getElementById("modalidade");
	combo.options[0].selected = "true";
//	var combo2 = document.getElementById("curso");
//	combo2.options[0].selected = "true";	
//	var combo3 = document.getElementById("etapa");
//	combo3.options[0].selected = "true";	
//	var combo4 = document.getElementById("turma");
//	combo4.options[0].selected = "true";	
//	var combo5 = document.getElementById("periodo");
//	combo5.options[0].selected = "true";		
    return false
  }  
  return true

}
function testa_tipo(tipo){
		if (tipo == "it") {
		document.getElementById('modalidade').disabled   = false;	    
		} else {
		document.getElementById('modalidade').disabled   = true;	    
		}
}
function testa_modalidade(modalidade){
		if (modalidade == "eg") {
		document.getElementById('etapa').disabled   = false;
		document.getElementById('turma').disabled   = true;	    			    
		} else {
		document.getElementById('etapa').disabled   = true;	    
		document.getElementById('turma').disabled   = true;	    		
		}
}

                         </script>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" background="../../../../img/fundo.gif" marginwidth="0" marginheight="0" onLoad="MM_preloadImages('../../../../img/menu_r1_c2_f3.gif','../../../../img/menu_r1_c2_f2.gif','../../../../img/menu_r1_c2_f4.gif','../../../../img/menu_r1_c4_f3.gif','../../../../img/menu_r1_c4_f2.gif','../../../../img/menu_r1_c4_f4.gif','../../../../img/menu_r1_c6_f3.gif','../../../../img/menu_r1_c6_f2.gif','../../../../img/menu_r1_c6_f4.gif','../../../../img/menu_r1_c8_f3.gif','../../../../img/menu_r1_c8_f2.gif','../../../../img/menu_r1_c8_f4.gif','../../../../img/menu_direita_r2_c1_f3.gif','../../../../img/menu_direita_r2_c1_f2.gif','../../../../img/menu_direita_r2_c1_f4.gif','../../../../img/menu_direita_r4_c1_f3.gif','../../../../img/menu_direita_r4_c1_f2.gif','../../../../img/menu_direita_r4_c1_f4.gif','../../../../img/menu_direita_r6_c1_f3.gif','../../../../img/menu_direita_r6_c1_f2.gif','../../../../img/menu_direita_r6_c1_f4.gif');">
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
      <%	call mensagens(4,9706,0,0) 
	  
	  
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
<FORM name="formulario" METHOD="POST" ACTION="../../../../relatorios/swd011.asp" onSubmit="return checksubmit()">                
        <table width="100%" border="0" align="right" cellpadding="0" cellspacing="0" class="tb_corpo">
          <tr class="tb_tit"> 
            <td width="653" height="15" class="tb_tit">Informe os crit&eacute;rios 
              para pesquisa </td>
          </tr>
          <tr> 
            <td valign="top"><table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="125" align="center" class="tb_subtit">Tipo</td>
                  <td width="125" align="center" class="tb_subtit">Modalidade</td>
                  <td width="125" class="tb_subtit"><div align="center">ETAPA </div></td>
                  <td width="125" class="tb_subtit"><div align="center">TURMA </div></td> 
                  <td width="500" class="tb_subtit"><div align="center">Per&iacute;odo</div></td>
                </tr>
                <tr>
                  <td width="125" align="center" class="form_dado_texto"><select name="tipo" id="tipo" class="select_style" onChange="testa_tipo(this.value);">
                    <option value="nulo" selected></option>
                    <option value="it" >Itens</option>
                    <option value="pr" >Projetos</option>
                  </select></td>
                  <td width="125" align="center" class="form_dado_texto"><select name="modalidade" id="modalidade" class="select_style" disabled  onChange="testa_modalidade(this.value);">
                    <option value="nulo" selected></option>
                    <option value="eg" >Em Geral</option>
                    <option value="pp" >Por Projetos</option>
                  </select></td>
                  <td width="125" align="center"><select name="etapa" class="select_style" onChange="recuperarTurma(this.value);">
                        <option value="999990" selected> </option>                  
                          <%		
unidade = 1
curso = 0
session("u_pub") = unidade
session("c_pub") = curso
		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&unidade&"AND CO_Curso='"&curso&"'"
		RS0b.Open SQL0b, CON0
		
		
While not RS0b.EOF
Etapa = RS0b("CO_Etapa")


		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&curso&"' AND CO_Etapa='"&Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
if Etapa=co_etapa then
%>
                          <option value="<%response.Write(Etapa)%>" selected> 
                          <%response.Write(NO_Etapa)%>
                          </option>
                          <%
else
%>
                          <option value="<%response.Write(Etapa)%>"> 
                          <%response.Write(NO_Etapa)%>
                          </option>
                          <%

end if
RS0b.MOVENEXT
WEND
%>
                        </select></td>
                  <td width="125"><div align="center">
                    <div id="divTurma">
                      <select name="turma" class="select_style" id="turma">
                        <option value="999990" selected> </option>
                      </select>
                    </div>
                  </div></td> 
                  <td width="500"><div align="center"><font class="form_dado_texto">
                    <select name="dia_de" id="dia_de" class="select_style">
                      <option value="1" selected>01</option>
                      <option value="2">02</option>
                      <option value="3">03</option>
                      <option value="4">04</option>
                      <option value="5">05</option>
                      <option value="6">06</option>
                      <option value="7">07</option>
                      <option value="8">08</option>
                      <option value="9">09</option>
                      <option value="10">10</option>
                      <option value="11">11</option>
                      <option value="12">12</option>
                      <option value="13">13</option>
                      <option value="14">14</option>
                      <option value="15">15</option>
                      <option value="16">16</option>
                      <option value="17">17</option>
                      <option value="18">18</option>
                      <option value="19">19</option>
                      <option value="20">20</option>
                      <option value="21">21</option>
                      <option value="22">22</option>
                      <option value="23">23</option>
                      <option value="24">24</option>
                      <option value="25">25</option>
                      <option value="26">26</option>
                      <option value="27">27</option>
                      <option value="28">28</option>
                      <option value="29">29</option>
                      <option value="30">30</option>
                      <option value="31">31</option>
                      </select>
                    /
                    <select name="mes_de" id="mes_de" class="select_style">
                      <option value="1" selected>janeiro</option>
                      <option value="2">fevereiro</option>
                      <option value="3">mar&ccedil;o</option>
                      <option value="4">abril</option>
                      <option value="5">maio</option>
                      <option value="6">junho</option>
                      <option value="7">julho</option>
                      <option value="8">agosto</option>
                      <option value="9">setembro</option>
                      <option value="10">outubro</option>
                      <option value="11">novembro</option>
                      <option value="12">dezembro</option>
                      </select>
                    /
                    <%response.write(ano_letivo)%>
                    <input name="ano_de" type="hidden" id="ano_de" value="<%response.write(ano_letivo)%>">
                    at&eacute;
                    <select name="dia_ate" id="dia_ate" class="select_style">
                      <% 
							 For i =1 to 31
							 dia=dia*1
							 if dia=i then 
								if dia<10 then
								dia="0"&dia
								end if
							 %>
                      <option value="<%response.Write(dia)%>" selected>
                        <%response.Write(dia)%>
                        </option>
                      <% else
								if i<10 then
								i="0"&i
								end if
							%>
                      <option value="<%response.Write(i)%>">
                        <%response.Write(i)%>
                        </option>
                      <% end if 
							next
							%>
                      </select>
                    /
                    <select name="mes_ate" id="mes_ate" class="select_style">
                      <%mes=mes*1
								if mes="1" or mes=1 then%>
                      <option value="1" selected>janeiro</option>
                      <% else%>
                      <option value="1">janeiro</option>
                      <%end if
								if mes="2" or mes=2 then%>
                      <option value="2" selected>fevereiro</option>
                      <% else%>
                      <option value="2">fevereiro</option>
                      <%end if
								if mes="3" or mes=3 then%>
                      <option value="3" selected>mar&ccedil;o</option>
                      <% else%>
                      <option value="3">mar&ccedil;o</option>
                      <%end if
								if mes="4" or mes=4 then%>
                      <option value="4" selected>abril</option>
                      <% else%>
                      <option value="4">abril</option>
                      <%end if
								if mes="5" or mes=5 then%>
                      <option value="5" selected>maio</option>
                      <% else%>
                      <option value="5">maio</option>
                      <%end if
								if mes="6" or mes=6 then%>
                      <option value="6" selected>junho</option>
                      <% else%>
                      <option value="6">junho</option>
                      <%end if
								if mes="7" or mes=7 then%>
                      <option value="7" selected>julho</option>
                      <% else%>
                      <option value="7">julho</option>
                      <%end if%>
                      <%if mes="8" or mes=8 then%>
                      <option value="8" selected>agosto</option>
                      <% else%>
                      <option value="8">agosto</option>
                      <%end if
								if mes="9" or mes=9 then%>
                      <option value="9" selected>setembro</option>
                      <% else%>
                      <option value="9">setembro</option>
                      <%end if
								if mes="10" or mes=10 then%>
                      <option value="10" selected>outubro</option>
                      <% else%>
                      <option value="10">outubro</option>
                      <%end if
								if mes="11" or mes=11 then%>
                      <option value="11" selected>novembro</option>
                      <% else%>
                      <option value="11">novembro</option>
                      <%end if
								if mes="12" or mes=12 then%>
                      <option value="12" selected>dezembro</option>
                      <% else%>
                      <option value="12">dezembro</option>
                      <%end if%>
                      </select>
                    /
                    <%response.write(ano_letivo)%>
                    </font>
                    <input name="ano_ate" type="hidden" id="ano_ate" value="<%response.write(ano_letivo)%>">
                  </div></td>
                </tr>
                <tr> 
                  <td colspan="5">&nbsp;</td>
                </tr>
                <tr> 
                  <td colspan="5"><hr width="1000"></td>
                </tr>
                <tr> 
                  <td colspan="5" valign="top">
                    <table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td width="33%">&nbsp;</td>
                          <td width="34%">&nbsp;</td>
                          <td width="33%" align="center"><input name="SUBMIT" type=SUBMIT class="botao_prosseguir" value="Prosseguir"></td>
                        </tr>
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