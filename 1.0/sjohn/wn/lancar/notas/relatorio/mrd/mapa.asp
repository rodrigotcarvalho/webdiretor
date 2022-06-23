<%On Error Resume Next%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes3.asp"-->

<%
opt=request.QueryString("opt")
obr = request.QueryString("obr")
imprime=request.QueryString("p")

if opt="ok" or imprime="1" or  opt= "vt" or opt="cln" then
dados= split(obr, "_" )
co_materia = dados(0)
unidades= dados(1)
grau= dados(2)
serie= dados(3)
turma= dados(4)
periodo = dados(5)
ano_letivo = dados(6)
co_prof = dados(7)
co_usr = session("co_usr")



else

co_materia = request.form("mat")
unidades= request.form("unidade")
grau= request.form("curso")
serie= request.form("etapa")
turma= request.form("turma")
periodo = request.form("periodo")
ano_letivo = request.form("ano_letivo")
co_prof = request.form("co_prof")
co_usr = session("co_usr")
end if
session("co_materia")=co_materia
session("unidades")=unidades
session("grau")=grau
session("serie")=serie
session("turma")=turma
session("periodo")=periodo

obr=unidades&"_"&grau&"_"&serie&"_"&turma&"_"&ano_letivo&"_"&co_materia


if serie = "999999" then
		response.redirect("tabelas2.asp?opt=err1&or=1&dd="&ano_letivo&"_"&grau&"_"&unidades&"_"&co_prof)
ELSEif co_materia = "999999" then
		response.redirect("tabelas2.asp?opt=err2&or=1&dd="&ano_letivo&"_"&grau&"_"&unidades&"_"&co_prof)
ELSEif turma = "999999" then
	response.redirect("tabelas2.asp?opt=err3&or=1&dd="&ano_letivo&"_"&grau&"_"&unidades&"_"&co_prof)

else


				Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
				Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidades &" AND CO_Curso = '"& grau &"' AND CO_Etapa = '"& serie &"' AND CO_Turma = '"& turma &"'"
		Set RS = CON.Execute(CONEXAO)

if RS.EOF then
response.Write("<div align=center><font size=2 face=Courier New, Courier, mono  color=#990000><b>Esta turma não está disponível no momento</b></font><br")
response.Write("<font size=2 face=Courier New, Courier, mono  color=#990000><a href=javascript:window.history.go(-1)>voltar</a></font></div>")

else
nota = RS("TP_Nota")
end if


%>
<html>
<head>
<title>Relat&oacute;rios</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<link rel="stylesheet" href="../estilos.css" type="text/css">
<script language="JavaScript" type="text/JavaScript">
<!--


function MM_popupMsg(msg) { //v1.0
  alert(msg);
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}
//-->
</script>
<script language="JavaScript" type="text/JavaScript">
<!--
function MM_openBrWindow(theURL,winName,features) { //v2.0
  window.open(theURL,winName,features);
}
//-->
</script>
</head>

<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<%
id0 = " > <a href='tabelas.asp?or=01&ano="&ano_letivo&"' class='linkum' target='_parent'>Selecionando a turma</a>"
id1 = " > Mapa de Resultados"

 call cabecalho(nivel) %>
<table width="1000" height="670" border="0" align="center" cellspacing="0" bgcolor="#FFFFFF">
  <tr> 
    <td valign="top"> <div align="center"> 
        <table width="1000" border="0" class="tb_caminho">
          <tr> 
            <td><font color="#FFFF33" size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="../inicio.asp" target="_parent" class="caminho">Web 
              Notas</a> 
              <%

	  response.Write(origem&id0&id1)
%>
              </font></td>
          </tr>
        </table>
        <%


			Set Conecta= Server.CreateObject("ADODB.Connection") 
		ABRIRgn = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		Conecta.Open ABRIRgn

call GeraNomes(co_materia,unidades,grau,serie,Conecta)

no_materia= session("no_materia")
no_unidades= session("no_unidades")
no_grau= session("no_grau")
no_serie= session("no_serie")


nome_prof = session("nome_prof") 
tp=	session("tp")

ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("m", now) 
data = dia &"/"& mes &"/"& ano
horario = hora & ":"& min
acesso_prof = session("acesso_prof")


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR
		
		Set RS = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE CO_Professor= "& co_prof &"AND NU_Unidade = "& unidades &" AND CO_Curso = '"& grau &"' AND CO_Etapa = '"& serie &"' AND CO_Turma = '"& turma &"' AND CO_Materia = '"& co_materia &"'"
		Set RS = CON.Execute(CONEXAO)

%>
        <br>
        <table width="1000" border="0" cellspacing="0">
          <tr> 
            <td width="219" valign="top"> <table width="100%" border="0" cellspacing="0">
                <%if opt = "ok" then%>
                <tr> 
                  <td> 
                    <%
		call mensagens(17,2,0)
%>
                  </td>
                </tr>
                <%end if
IF imprime="1"then%>
                <tr> 
                  <td> 
                    <%	call mensagens(4,0,0) 

%>
                  </td>
                </tr>
                <tr> 
                  <td> 
                    <%	call mensagens(6,2,0) 

%>
                  </td>
                </tr>
                <% else%>
                <tr> 
                  <td> 
                    <%	call mensagens(19,0,0) 

%>
                  </td>
                </tr>
                <tr> 
                  <td> 
                    <%	call mensagens(52,2,0) 

%>
                  </td>
                </tr>
                <% end if%>
                <tr> 
                  <td> </td>
                </tr>
              </table></td>
            <td width="770" valign="top"> <div align="right"> 
                <%
Call contalunos(CAMINHOa,unidades,grau,serie,turma)

if nota ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na
Call mapa_notas(CAMINHOa,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,1,0)
elseif nota="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb
	Call mapa_notas(CAMINHOa,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,2,0)
elseif nota ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc
		Call mapa_notas(CAMINHOa,CAMINHOn,unidades,grau,serie,turma,co_materia,periodo,ano_letivo,co_usr,3,0)
else
		response.Write("ERRO")
End if

 %>
              </div></td>
          </tr>
        </table>
        <%	
    Set RS = Nothing
	    Set RS2 = Nothing
		    Set RS3 = Nothing
End if
%>
        <% IF imprime="1"then %>
        <table width="1000" border="0" cellspacing="0">
          <tr> 
            <td width="219">&nbsp;</td>
            <td width="770" class="tb_voltar"><font color="#669999" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="mapa.asp?or=02&ano=<%=ano_letivo%>&opt=vt" target="_parent" class="voltar1">&lt; 
              Sair da versão de impressão</a></strong></font></td>
          </tr>
        </table>
        <%else%>
        <table width="1000" border="0" cellspacing="0">
          <tr> 
            <td width="219">&nbsp;</td>
            <td width="770" class="tb_voltar"><font color="#669999" size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="tabelas.asp?or=02&ano=<%=ano_letivo%>" target="_parent" class="voltar1">&lt; 
              Voltar para Relat&oacute;rios</a></strong></font></td>
          </tr>
        </table>
        <%end if %>
      </div></td>
  </tr>
  <tr>
    <td height="40" valign="top"><img src="../img/rodape.jpg" width="1000" height="40"></td>
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
response.redirect("../inc/erro.asp")
end if
%>