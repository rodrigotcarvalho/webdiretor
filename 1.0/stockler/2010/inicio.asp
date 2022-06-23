<%On Error Resume Next%>

<!--#include file="inc/funcoes.asp"-->
<%
nivel=0
permissao = session("permissao")
ano_letivo = session("ano_letivo") 
ano_info=nivel&"_0_"&ano_letivo
chave="WR"
session("sistema_local")="WR"

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

 call navegacao (CON,chave,nivel)
navega=Session("caminho")
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
//-->
</script>
<link href="estilos.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0" background="img/fundo.gif" onLoad="<%response.Write(SESSION("onLoad"))%>">
<%
call cabecalho(nivel)
%>
        
<table width="1000" height="650" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
            
    <td height="14" class="front1"> 
      <table width="100%" height="83" border="0" class="front1">
        <tr> 
                  <td height="70" class="front1">&#8226; 
                SEJA BEM-VINDO ! Obrigado por estar utilizando esta nova ferramenta 
                de apoio a administra&ccedil;&atilde;o acad&ecirc;mica de nossa 
                escola. Estamos constantemente procurando otimizar os nossos procedimentos 
                operacionais e esperamos prover maiores benef&iacute;cios aos 
                nossos alunos , respons&aacute;veis e tamb&eacute;m ao corpo docente. 
                Por isso contamos com a sua colabora&ccedil;&atilde;o nesta fase 
                de implementa&ccedil;&atilde;o.</td>
                </tr>
              </table>
			 
            </td>
          </tr>
          <tr>
            
    <td height="160" class="front1"> 
      <table width="100%" height="160" border="0" class="front2">
        <tr> 
                  <td class="front2">&#8226; 
              ACESSO AO SISTEMA ! A escola forneceu a voc&ecirc; uma &#8220;Chave&#8221; 
              para o acesso ao sistema. Esta chave que &eacute; composta de Usu&aacute;rio 
              e Senha o identificar&aacute; sempre que utilizar as fun&ccedil;&otilde;es 
              do sistema. Assim, ao realizar o seu primeiro acesso, procure alterar 
              essa &#8220;Chave&#8221; a fim de personalizar a sua identifica&ccedil;&atilde;o 
              . A fun&ccedil;&atilde;o Alterar Senha est&aacute; localizada na 
              parte superior da tela Web Diretor e poder&aacute; ser acionada 
              caso necessite modificar o seu c&oacute;digo usu&aacute;rio, sua 
              senha ou seu email de contato.Recomendamos que a senha seja formada 
              por n&uacute;meros e letras (Exemplos: HKR55089 , 3700SLL4 , L4R8J3552). 
              Lembre-se tamb&eacute;m de verificar logo abaixo do cabe&ccedil;alho 
              no canto esquerdo da tela se seu nome est&aacute; aparecendo corretamente 
              com a informa&ccedil;&atilde;o do dia e hora de seu &uacute;ltimo 
              acesso ( somente na primeira conex&atilde;o estas &uacute;ltimas 
              informa&ccedil;&otilde;es n&atilde;o ser&atilde;o exibidas). Veja 
              ainda que ao lado direito ser&aacute; exibido o dia da semana e 
              o dia , m&ecirc;s e ano que est&aacute; se conectando ao sistema. 
              Caso tenha qualquer coment&aacute;rio ou d&uacute;vida sobre algum 
              procedimento operacional ou queira reportar uma falha do sistema, 
              utilize a fun&ccedil;&atilde;o Fale Conosco onde uma equipe t&eacute;cnica 
              examinar&aacute; e responder&aacute; suas solicita&ccedil;&otilde;es.</td>
                </tr>
              </table>
            
    </td>
          </tr>
          <tr> 
            
    <td height="150"class="front1" >
<table width="100%" height="150" border="0">
        <tr> 
          <td class="front1"> 
            &#8226; 
              ANO LETIVO E SISTEMAS AUTORIZADOS ! Ao se conectar ao Web Diretor 
              , duas &#8220;Listas&#8221; ficar&atilde;o dispon&iacute;veis para 
              sua utiliza&ccedil;&atilde;o. A primeira trata-se do &#8220;Ano 
              Letivo&#8221; onde estar&aacute; sempre apontado o Ano vigente do 
              estabelecimento de ensino. Isso indicar&aacute; que todas as opera&ccedil;&otilde;es 
              realizadas no sistema ser&atilde;o gravadas utilizando-se como refer&ecirc;ncia 
              este ano. Se desejarmos por exemplo registrar uma avalia&ccedil;&atilde;o 
              de um aluno neste ano , suas notas refletir&atilde;o o seu desempenho 
              escolar desse ano. Caso seja necess&aacute;rio a qualquer momento 
              rever algum dado passado, de anos anteriores, basta fazer esta escolha 
              alterando o campo ano letivo. Mas lembre-se que provavelmente o 
              seu acesso s&oacute; ser&aacute; permitido apenas para consultas 
              . O sistema far&aacute; restri&ccedil;&otilde;es para permitir altera&ccedil;&otilde;es 
              de dados nos anos anteriores. A pr&oacute;xima lista exibe os &#8220;Sistemas 
              Autorizados&#8221; onde ser&aacute; habilitada a rela&ccedil;&atilde;o 
              dos sistemas que sua chave de acesso permitir. Ao acessar um sistema 
              autorizado aparecer&atilde;o os M&oacute;dulos e suas respectivas 
              fun&ccedil;&otilde;es. Algumas fun&ccedil;&otilde;es poder&atilde;o 
              ter seu acesso apenas para consulta , neste caso n&atilde;o podendo 
              alterar qualquer dado exibido na tela.</td>
        </tr>
      </table> 
    </td>
          </tr>
          <tr valign="bottom" bgcolor="#FFFFFF"> 
            <td background="img/fundo_inicio.gif">
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