
<html xmlns:v="urn:schemas-microsoft-com:vml"
xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:w="urn:schemas-microsoft-com:office:word"
xmlns="http://www.w3.org/TR/REC-html40">

<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=ProgId content=Word.Document>
<meta name=Generator content="Microsoft Word 10">
<meta name=Originator content="Microsoft Word 10">
<link rel=File-List href="Bloqueto_arquivos/filelist.xml">
<link rel=Edit-Time-Data href="Bloqueto_arquivos/editdata.mso">
<!--[if !mso]>
<style>
v\:* {behavior:url(#default#VML);}
o\:* {behavior:url(#default#VML);}
w\:* {behavior:url(#default#VML);}
.shape {behavior:url(#default#VML);}
</style>
<![endif]-->
<title></title>
<!--[if gte mso 9]><xml>
 <o:DocumentProperties>
  <o:Author>Rodrigo Tovar de Carvalho</o:Author>
  <o:LastAuthor>Rodrigo Tovar de Carvalho</o:LastAuthor>
  <o:Revision>5</o:Revision>
  <o:TotalTime>130</o:TotalTime>
  <o:Created>2003-10-21T03:53:00Z</o:Created>
  <o:LastSaved>2003-11-01T19:52:00Z</o:LastSaved>
  <o:Pages>1</o:Pages>
  <o:Words>289</o:Words>
  <o:Characters>1566</o:Characters>
  <o:Company>RoDan Tecnologia da Informação ltda</o:Company>
  <o:Lines>13</o:Lines>
  <o:Paragraphs>3</o:Paragraphs>
  <o:CharactersWithSpaces>1852</o:CharactersWithSpaces>
  <o:Version>10.2625</o:Version>
 </o:DocumentProperties>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <w:WordDocument>
  <w:HyphenationZone>21</w:HyphenationZone>
  <w:Compatibility>
   <w:BreakWrappedTables/>
   <w:SnapToGridInCell/>
   <w:WrapTextWithPunct/>
   <w:UseAsianBreakRules/>
  </w:Compatibility>
  <w:BrowserLevel>MicrosoftInternetExplorer4</w:BrowserLevel>
 </w:WordDocument>
</xml><![endif]-->
<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:Verdana;
	panose-1:2 11 6 4 3 5 4 4 2 4;
	mso-font-charset:0;
	mso-generic-font-family:swiss;
	mso-font-pitch:variable;
	mso-font-signature:647 0 0 0 159 0;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{mso-style-parent:"";
	margin:0cm;
	margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:12.0pt;
	font-family:"Times New Roman";
	mso-fareast-font-family:"Times New Roman";}
@page Section1
	{size:612.0pt 792.0pt;
	margin:215.65pt 72.0pt 9.0pt 81.0pt;
	mso-header-margin:35.4pt;
	mso-footer-margin:35.4pt;
	mso-paper-source:0;}
div.Section1
	{page:Section1;}
-->
</style>
<!--[if gte mso 10]>
<style>
 /* Style Definitions */
 table.MsoNormalTable
	{mso-style-name:"Tabela normal";
	mso-tstyle-rowband-size:0;
	mso-tstyle-colband-size:0;
	mso-style-noshow:yes;
	mso-style-parent:"";
	mso-padding-alt:0cm 5.4pt 0cm 5.4pt;
	mso-para-margin:0cm;
	mso-para-margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	font-family:"Times New Roman";}
table.MsoTableGrid
	{mso-style-name:"Tabela com grade";
	mso-tstyle-rowband-size:0;
	mso-tstyle-colband-size:0;
	border:solid windowtext 1.0pt;
	mso-border-alt:solid windowtext .5pt;
	mso-padding-alt:0cm 5.4pt 0cm 5.4pt;
	mso-border-insideh:.5pt solid windowtext;
	mso-border-insidev:.5pt solid windowtext;
	mso-para-margin:0cm;
	mso-para-margin-bottom:.0001pt;
	mso-pagination:widow-orphan;
	font-size:10.0pt;
	font-family:"Times New Roman";}
</style>
<![endif]--><!--[if gte mso 9]><xml>
 <o:shapedefaults v:ext="edit" spidmax="4098"/>
</xml><![endif]--><!--[if gte mso 9]><xml>
 <o:shapelayout v:ext="edit">
  <o:idmap v:ext="edit" data="1"/>
 </o:shapelayout></xml><![endif]-->
</head>
                             <%
matricula = request.querystring("codigo")
venc=request.querystring("opt")

Set CON2 = Server.CreateObject("ADODB.Connection") 
CAMINHO2 = "e:\home\bretanha\dados\Alunos.mdb"
ABRIR2 = "DBQ="& CAMINHO2 & ";Driver={Microsoft Access Driver (*.mdb)}"
CON2.Open ABRIR2
Set RS2 = Server.CreateObject("ADODB.Recordset")
CONEXAO2 = "SELECT * FROM Alunos WHERE CO_Matricula_Escola = " & matricula 
RS2.Open CONEXAO2, CON2

Set CON3 = Server.CreateObject("ADODB.Connection") 
CAMINHO3 = "e:\home\bretanha\dados\Bloqueto.mdb"
ABRIR3 = "DBQ="& CAMINHO3 & ";Driver={Microsoft Access Driver (*.mdb)}"
CON3.Open ABRIR3
Set rsblo = Server.CreateObject("ADODB.Recordset")
CONEXAO3 = "SELECT * FROM bloqueto WHERE CO_Matricula_Escola=" & matricula &"AND NU_Cota = " & venc 
rsblo.Open CONEXAO3, CON3
%>
<body lang=PT-BR style='tab-interval:35.4pt'>

<div class=Section1>

<p class=MsoNormal align=center style='text-align:center'><span
style='font-size:10.0pt;font-family:Verdana;color:#000066'>Para imprimir o
boleto clique em <b>Arquivo</b> e <b>Imprimir</b>, ou <b>File</b> e <b>Print</b>,
no menu.<br>
Usar papel branco de gramatura mínima de 50 g/m<sup>2</sup>, com impressão
preta ou azul.</span><span style='color:#000066'><o:p></o:p></span></p>

<p class=MsoNormal><o:p>&nbsp;</o:p></p>

<div align=center>

    <table class=MsoNormalTable border=1 cellspacing=0 cellpadding=0 width=661
 style='width:495.45pt;margin-left:-5.5pt;border-collapse:collapse;border:none;
 mso-border-top-alt:solid windowtext .5pt;mso-padding-alt:0cm 3.5pt 0cm 3.5pt'>
      <tr style='mso-yfti-irow:0;height:9.75pt'> 
        <td colspan=2 valign=top style='width:93.5pt;border:none;
  padding:0cm 3.5pt 0cm 3.5pt;height:9.75pt'> <p class=MsoNormal><o:p>&nbsp;</o:p></p></td>
        <td colspan=3 style='width:90.8pt;border:none;border-bottom:solid windowtext 1.0pt;
  mso-border-bottom-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:9.75pt'> <p class=MsoNormal style='text-indent:39.5pt'><span style='font-size:10.0pt;
  font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td colspan=11 valign=bottom style='width:311.15pt;border:none;
  border-bottom:solid windowtext 1.0pt;mso-border-bottom-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:9.75pt'> <p class=MsoNormal align=right style='text-align:right'><b style='mso-bidi-font-weight:
  normal'><span style='font-size:8.0pt;font-family:Arial'>Recibo do Sacado</span></b><span
  style='font-size:10.0pt;font-family:Arial'><o:p></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:1;height:12.0pt'> 
        <td colspan=2 rowspan=4 valign=top style='width:93.5pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:12.0pt'> <p class=MsoNormal align=center style='text-align:center'> 
            <!--[if gte vml 1]><v:shapetype
   id="_x0000_t75" coordsize="21600,21600" o:spt="75" o:preferrelative="t"
   path="m@4@5l@4@11@9@11@9@5xe" filled="f" stroked="f">
   <v:stroke joinstyle="miter"/>
   <v:formulas>
    <v:f eqn="if lineDrawn pixelLineWidth 0"/>
    <v:f eqn="sum @0 1 0"/>
    <v:f eqn="sum 0 0 @1"/>
    <v:f eqn="prod @2 1 2"/>
    <v:f eqn="prod @3 21600 pixelWidth"/>
    <v:f eqn="prod @3 21600 pixelHeight"/>
    <v:f eqn="sum @0 0 1"/>
    <v:f eqn="prod @6 1 2"/>
    <v:f eqn="prod @7 21600 pixelWidth"/>
    <v:f eqn="sum @8 21600 0"/>
    <v:f eqn="prod @7 21600 pixelHeight"/>
    <v:f eqn="sum @10 21600 0"/>
   </v:formulas>
   <v:path o:extrusionok="f" gradientshapeok="t" o:connecttype="rect"/>
   <o:lock v:ext="edit" aspectratio="t"/>
  </v:shapetype><v:shape id="_x0000_i1025" type="#_x0000_t75" style='width:43.5pt;
   height:43.5pt'>
   <v:imagedata src="Bloqueto_arquivos/image001.png" o:title="logo_boleto"/>
  </v:shape><![endif]-->
            <![if !vml]>
            <img width=58 height=58
  src="Bloqueto_arquivos/image002.jpg" v:shapes="_x0000_i1025"> 
            <![endif]>
          </p></td>
        <td width=53 style='width:39.6pt;border:none;border-bottom:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:12.0pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Aluno:<o:p></o:p></span></p></td>
        <td colspan=13 style='width:362.35pt;border:solid windowtext 1.0pt;
  border-left:none;mso-border-top-alt:solid windowtext .5pt;mso-border-bottom-alt:
  solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;padding:
  0cm 3.5pt 0cm 3.5pt;height:12.0pt'> <p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
  style='font-size:8.0pt;font-family:Arial'><o:p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (RS2("NO_Aluno"))%>
            </font></o:p></span></b></p></td>
      </tr>
      <tr style='mso-yfti-irow:2;height:7.1pt'> 
        <td colspan=9 valign=bottom style='width:241.35pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:7.1pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Matrícula:<o:p></o:p></span></p></td>
        <td colspan=5 valign=bottom style='width:160.6pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:7.1pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Vencimento:<o:p></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:3;height:15.75pt'> 
        <td colspan=9 style='width:241.35pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:15.75pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (RS2("CO_Matricula_Escola"))%>
            </font></o:p></span></p></td>
        <td colspan=5 style='width:160.6pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:15.75pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (rsblo("DA_Vencimento"))%>
            </font>&nbsp;</o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:4;height:8.7pt'> 
        <td colspan=5 valign=top style='width:147.0pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:8.7pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>N. 
            Cota:<o:p></o:p></span></p></td>
        <td colspan=4 valign=top style='width:94.35pt;border-top:solid windowtext 1.0pt;
  border-left:none;border-bottom:none;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:8.7pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>N.Carnê 
            de Pagto:<o:p></o:p></span></p></td>
        <td colspan=3 valign=top style='width:92.9pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:8.7pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Nosso 
            Número:<o:p></o:p></span></p></td>
        <td colspan=2 valign=top style='width:67.7pt;border:none;border-right:
  solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:
  solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:
  solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;padding:
  0cm 3.5pt 0cm 3.5pt;height:8.7pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Valor 
            Cobrado:<o:p></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:5;height:6.4pt'> 
        <td colspan=2 valign=top style='width:93.5pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.4pt'> <p align="center" class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
  style='font-size:6.5pt;font-family:Arial'>CNPJ: 34.156.620/0001-36<o:p></o:p></span></b></p></td>
        <td colspan=5 style='width:147.0pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.4pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (rsblo("NU_Cota"))%>
            </font>&nbsp;</o:p></span></p></td>
        <td colspan=4 style='width:94.35pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.4pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (rsblo("NU_Bloqueto"))%>
            </font>&nbsp;</o:p></span></p></td>
        <td colspan=3 style='width:92.9pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.4pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (rsblo("CO_Nosso_Numero"))%>
            </font></o:p></span></p></td>
        <td colspan=2 style='width:67.7pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.4pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (FormatCurrency(rsblo("VA_Inicial")))%>
            </font></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:6;height:4.15pt'> 
        <td colspan=10 rowspan=2 valign=top style='width:318.05pt;
  border:none;padding:0cm 3.5pt 0cm 3.5pt;height:4.15pt'> <p class=MsoNormal align=right style='text-align:right'><span
  style='font-size:7.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td colspan=2 valign=bottom style='width:49.85pt;border:none;
  border-bottom:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-bottom-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:4.15pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td colspan=3 rowspan=2 style='width:90.65pt;border:none;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:4.15pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;font-family:Arial'>Autenticação Mecânica<o:p></o:p></span></p></td>
        <td width=49 valign=bottom style='width:36.9pt;border:none;border-bottom:
  solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;mso-border-top-alt:
  solid windowtext .5pt;mso-border-bottom-alt:solid windowtext .5pt;padding:
  0cm 3.5pt 0cm 3.5pt;height:4.15pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:7;height:4.1pt'> 
        <td colspan=2 valign=bottom style='width:49.85pt;border:none;
  border-left:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:4.1pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td width=49 valign=bottom style='width:36.9pt;border:none;mso-border-top-alt:
  solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:4.1pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:8;height:28.9pt'> 
        <td colspan=16 valign=bottom style='width:495.45pt;border:none;
  padding:0cm 3.5pt 0cm 3.5pt;height:28.9pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:22.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:9;height:4.0pt'> 
        <td colspan=16 valign=bottom style='width:495.45pt;border:none;
  padding:0cm 3.5pt 0cm 3.5pt;height:4.0pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'>.......................................................................................................................................................................................................................<o:p></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:10;height:10.15pt'> 
        <td colspan=2 valign=top style='width:93.5pt;border:none;
  padding:0cm 3.5pt 0cm 3.5pt;height:10.15pt'> <p class=MsoNormal><img width=109 height=29 id="_x0000_i1026"
  src="Bloqueto_arquivos/bradesco.gif"></p></td>
        <td colspan=14 valign=bottom style='width:401.95pt;border:none;
  padding:0cm 3.5pt 0cm 3.5pt;height:10.15pt'> <p class=MsoNormal><span style='font-family:Verdana;mso-bidi-font-family:
  Arial'>|</span><b style='mso-bidi-font-weight:normal'><span style='font-family:
  Arial'>237-2</span></b><span style='font-family:Verdana;mso-bidi-font-family:
  Arial'>|</span><span style='font-size:12.0pt;font-family:Arial'><o:p> 
            <%response.write (rsblo("CO_Superior"))%>
            </o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:11;height:10.1pt'> 
        <td colspan=2 style='width:93.5pt;border:none;border-bottom:solid windowtext 1.0pt;
  mso-border-bottom-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:10.1pt'> <p class=MsoNormal><o:p>&nbsp;</o:p></p></td>
        <td colspan=3 style='border:none;border-bottom:solid windowtext 1.0pt;
  mso-border-bottom-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:10.1pt'> <p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
  style='font-family:Verdana;mso-bidi-font-family:Arial'><o:p>&nbsp;</o:p></span></b></p></td>
        <td colspan=11 style='width:311.15pt;border:none;border-bottom:
  solid windowtext 1.0pt;mso-border-bottom-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:10.1pt'> <p class=MsoNormal align=right style='text-align:right'><b style='mso-bidi-font-weight:
  normal'><span style='font-size:8.0pt;font-family:Arial'>Ficha de Compensação<o:p></o:p></span></b></p></td>
      </tr>
      <tr style='mso-yfti-irow:12;height:7.4pt'> 
        <td colspan=2 valign=top style='width:93.5pt;border-top:none;
  border-left:solid windowtext 1.0pt;border-bottom:solid windowtext 1.0pt;
  border-right:none;mso-border-top-alt:solid windowtext .5pt;mso-border-top-alt:
  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-bottom-alt:
  solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:7.4pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Local 
            de Pagamento<o:p></o:p></span></p></td>
        <td colspan=11 style='width:287.35pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:7.4pt'> <p class=MsoNormal><b style='mso-bidi-font-weight:normal'><span
  style='font-size:8.0pt;font-family:Arial'>Pagável Preferencialmente nas agências 
            do Bradesco<o:p></o:p></span></b></p></td>
        <td width=63 valign=top style='width:46.9pt;border:none;border-bottom:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:7.4pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Vencimento<o:p></o:p></span></p></td>
        <td colspan=2 style='width:67.7pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:7.4pt'> <p class=MsoNormal><span style='font-size:8.0pt;font-family:Verdana;
  mso-bidi-font-family:Arial'><o:p><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (rsblo("DA_Vencimento"))%>
            </font></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:13;height:6.25pt'> 
        <td colspan=2 rowspan=2 valign=top style='width:93.5pt;border-top:
  none;border-left:solid windowtext 1.0pt;border-bottom:solid windowtext 1.0pt;
  border-right:none;mso-border-top-alt:solid windowtext .5pt;mso-border-top-alt:
  solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;mso-border-bottom-alt:
  solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Cedente<o:p></o:p></span></p></td>
        <td colspan=11 rowspan=2 style='width:287.35pt;border-top:none;
  border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:8.0pt;font-family:Arial'><o:p><font size="2" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (rsblo("NO_Cedente"))%>
            </font></o:p></span></p></td>
        <td colspan=3 valign=top style='width:114.6pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Agência 
            / Código<span style='mso-spacerun:yes'>  </span>Cedente<o:p></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:14;height:6.25pt'> 
        <td colspan=3 valign=top style='width:114.6pt;border-top:none;
  border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (rsblo("CO_Agencia"))%>
            </font>/<font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (rsblo("CO_Conta"))%>
            </font>&nbsp;</o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:15;height:6.25pt'> 
        <td colspan=2 valign=top style='width:93.5pt;border-top:none;
  border-left:solid windowtext 1.0pt;border-bottom:none;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Data 
            do Documento<o:p></o:p></span></p></td>
        <td colspan=4 style='width:91.25pt;border:none;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-right-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>N. 
            Documento<o:p></o:p></span></p></td>
        <td colspan=2 style='width:83.6pt;border:none;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-right-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Espécie 
            Doc.<o:p></o:p></span></p></td>
        <td colspan=2 style='width:49.7pt;border:none;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-right-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:6.25pt'> <p class=MsoNormal style='tab-stops:-345.5pt'><span style='font-size:7.0pt;
  font-family:Arial'>Aceite<o:p></o:p></span></p></td>
        <td colspan=3 style='width:62.8pt;border:none;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-right-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Data 
            Processam.<o:p></o:p></span></p></td>
        <td colspan=3 valign=top style='width:114.6pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Nosso 
            Número<o:p></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:16;height:6.25pt'> 
        <td colspan=2 valign=top style='width:93.5pt;border:solid windowtext 1.0pt;
  border-top:none;mso-border-left-alt:solid windowtext .5pt;mso-border-bottom-alt:
  solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;padding:
  0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (rsblo("DA_Processamento"))%>
            </font><o:p>&nbsp;</o:p></span></p></td>
        <td colspan=4 style='width:91.25pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p>&nbsp;</o:p><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (rsblo("NU_Cota"))%>
            </font></span></p></td>
        <td colspan=2 style='width:83.6pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td colspan=2 style='width:49.7pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td colspan=3 style='width:62.8pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (rsblo("DA_Processamento"))%>
            </font></o:p></span></p></td>
        <td colspan=3 valign=top style='width:114.6pt;border-top:none;
  border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p>&nbsp;006/<font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (rsblo("CO_Nosso_Numero"))%>
            </font></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:17;height:6.25pt'> 
        <td colspan=2 valign=top style='width:93.5pt;border-top:none;
  border-left:solid windowtext 1.0pt;border-bottom:none;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;font-family:Arial'>Uso Banco<o:p></o:p></span></p></td>
        <td colspan=2 style='width:53.0pt;border:none;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-right-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;font-family:Arial'>Carteira<o:p></o:p></span></p></td>
        <td colspan=2 style='width:38.25pt;border:none;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-right-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:7.0pt;font-family:Arial'>Espécie<o:p></o:p></span></p></td>
        <td colspan=2 style='width:83.6pt;border:none;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-right-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Quantidade<o:p></o:p></span></p></td>
        <td colspan=5 style='width:112.5pt;border:none;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-right-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Valor<o:p></o:p></span></p></td>
        <td colspan=3 valign=top style='width:114.6pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>(=) 
            Valor do Documento<o:p></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:18;height:14.5pt'> 
        <td colspan=2 style='width:93.5pt;border:solid windowtext 1.0pt;
  border-top:none;mso-border-left-alt:solid windowtext .5pt;mso-border-bottom-alt:
  solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;padding:
  0cm 3.5pt 0cm 3.5pt;height:14.5pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td colspan=2 style='width:53.0pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:14.5pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'>006<o:p></o:p></span></p></td>
        <td colspan=2 style='width:38.25pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:14.5pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'>R$<o:p></o:p></span></p></td>
        <td colspan=2 style='width:83.6pt;border-top:none;border-left:none;
  border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:14.5pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td colspan=5 style='width:112.5pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:14.5pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td colspan=3 valign=bottom style='width:114.6pt;border-top:none;
  border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:14.5pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
            <%response.write (FormatCurrency(rsblo("VA_Inicial")))%>
            </font></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:19;height:6.25pt'> 
        <td colspan=13 valign=bottom style='width:380.85pt;border-top:none;
  border-left:solid windowtext 1.0pt;border-bottom:none;border-right:solid windowtext 1.0pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Instruções 
            (todas as informações deste bloqueto são de exclusiva responsabilidade 
            do cedente)<o:p></o:p></span></p></td>
        <td colspan=3 valign=top style='width:114.6pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>(-) 
            Desconto<o:p></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:20;height:12.0pt'> 
        <td colspan=13 style='width:380.85pt;border-top:none;border-left:
  solid windowtext 1.0pt;border-bottom:none;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:12.0pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td colspan=3 valign=top style='width:114.6pt;border-top:none;
  border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:12.0pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:10.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:21;height:6.25pt'> 
        <td colspan=13 valign=top style='width:380.85pt;border-top:none;
  border-left:solid windowtext 1.0pt;border-bottom:none;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:8.0pt;font-family:Arial'><o:p><strong> 
            <%response.write (rsblo("TX_Msg_01"))%>
            <br>
            </strong></o:p></span></p></td>
        <td colspan=3 valign=top style='width:114.6pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>(-) 
            Outras deduções (abatimento)<o:p></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:22;height:8.3pt'> 
        <td colspan=13 valign=bottom style='width:380.85pt;border-top:none;
  border-left:solid windowtext 1.0pt;border-bottom:none;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:8.3pt'> <p class=MsoNormal><span style='font-size:8.0pt;font-family:Arial'><strong> 
            <%response.write (rsblo("TX_Msg_02"))%>
            </strong></span><span style='font-size:7.0pt;font-family:Arial'><o:p></o:p></span></p></td>
        <td colspan=3 valign=top style='width:114.6pt;border-top:none;
  border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:8.3pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:10.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:23;height:6.25pt'> 
        <td colspan=13 valign=top style='width:380.85pt;border-top:none;
  border-left:solid windowtext 1.0pt;border-bottom:none;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Após 
            o vencimento pagar somente no Bradesco com multa de 2% + permanência 
            diária de 0,05%<o:p></o:p></span><span style='font-size:9.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td colspan=3 valign=top style='width:114.6pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>(+) 
            Mora / Multa (juros)<o:p></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:24;height:6.25pt'> 
        <td colspan=13 valign=bottom style='width:380.85pt;border-top:none;
  border-left:solid windowtext 1.0pt;border-bottom:none;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span
  style='font-size:7.0pt;font-family:Arial'><o:p><span style='font-size:7.0pt;font-family:Arial'><b style='mso-bidi-font-weight:normal'><span
  style='font-size:7.0pt;font-family:Arial'>ATEN&Ccedil;&Atilde;O!</span></b></span></o:p></span></p></td>
        <td colspan=3 valign=top style='width:114.6pt;border-top:none;
  border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:10.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:25;height:6.25pt'> 
        <td colspan=13 rowspan=3 valign="top" style='width:380.85pt;border-top:none;
  border-left:solid windowtext 1.0pt;border-bottom:none;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'> 
            </span><span style='font-size:7.0pt;font-family:Arial'><span
  style='font-size:7.0pt;font-family:Arial'><o:p></o:p></span>PAGAMENTO após 30 
            dias do vencimento SOMENTE NA ESCOLA<o:p><br>
            <span style='font-size:8.0pt;font-family:Arial'><strong>
            <%response.write (rsblo("TX_Msg_08"))%>
            </strong></span> </o:p></span></p>
        </td>
        <td colspan=3 valign=top style='width:114.6pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>(+) 
            Outros acréscimos<o:p></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:26;height:6.25pt'> 
        <td colspan=3 valign=top style='width:114.6pt;border-top:none;
  border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:10.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:27;height:6.25pt'> 
        <td colspan=3 valign=top style='width:114.6pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>(=) 
            Valor cobrado<o:p></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:28;height:6.25pt'> 
        <td width=117 valign=top style='width:86.4pt;border-top:none;border-left:
  solid windowtext 1.0pt;border-bottom:solid windowtext 1.0pt;border-right:
  none;mso-border-left-alt:solid windowtext .5pt;mso-border-bottom-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td colspan=5 style='width:98.35pt;border:none;border-bottom:solid windowtext 1.0pt;
  mso-border-bottom-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td width=98 style='width:55.75pt;border:none;border-bottom:solid windowtext 1.0pt;
  mso-border-bottom-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td colspan=2 style='width:43.0pt;border:none;border-bottom:solid windowtext 1.0pt;
  mso-border-bottom-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:8.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td colspan=4 style='width:97.35pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:9.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
        <td colspan=3 valign=top style='width:114.6pt;border-top:none;
  border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-left-alt:solid windowtext .5pt;mso-border-left-alt:solid windowtext .5pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal align=center style='text-align:center'><span
  style='font-size:10.0pt;font-family:Arial'><o:p>&nbsp;</o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:29;height:2.25pt'> 
        <td colspan=10 valign=bottom style='width:318.05pt;border:none;
  border-left:solid windowtext 1.0pt;mso-border-left-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:2.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Sacado:<o:p></o:p></span></p></td>
        <td colspan=6 valign=bottom style='width:177.4pt;border:none;
  border-right:solid windowtext 1.0pt;mso-border-top-alt:solid windowtext .5pt;
  mso-border-top-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:2.25pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Matrícula 
            / Aluno<o:p></o:p></span></p></td>
      </tr>
      <tr style='mso-yfti-irow:30;height:2.25pt'> 
        <td colspan=10 valign=top style='width:300.4pt;border-top:none;border-left:
  solid windowtext 1.0pt;border-bottom:solid windowtext 1.0pt;border-right:
  none;mso-border-left-alt:solid windowtext .5pt;mso-border-bottom-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:6.25pt'> <p class=MsoNormal><span style='font-size:8.0pt;font-family:Arial'><o:p><strong><font size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
            <%response.write (rsblo("NO_Responsavel"))%>
            </strong></font></strong></o:p></span></p>
          <p class=MsoNormal><span style='font-size:8.0pt;font-family:Arial'><o:p> 
            <%response.write (rsblo("NO_Logradouro_Empresa"))%>
            , 
            <%response.write (rsblo("NU_Logradouro_Empresa"))%>
            <%If len(rsblo("TX_Complemento_Logradouro_Empresa")) <> 0  Then %>
            - 
            <% response.write (rsblo("TX_Complemento_Logradouro_Empresa"))
												End IF%>
            </o:p></span></p>
          <p class=MsoNormal><span style='font-size:8.0pt;font-family:Arial'><o:p> 
            <%response.write (rsblo("NO_Bairro_Empresa"))%>
            <%If len(rsblo("NO_Bairro_Empresa")) <> 0  Then %>
            - 
            <% END IF%>
            <%response.write (rsblo("NO_Cidade_Empresa"))%>
            - 
            <%response.write (rsblo("SG_UF_Empresa"))%>
            </o:p></span></p>
          <p class=MsoNormal><span style='font-size:8.0pt;font-family:Arial'><o:p> 
            <%response.write (rsblo("CO_CEP_Empresa"))%>
            </o:p></span></p></td>
        <td colspan=6 valign=top style='width:177.4pt;border-top:none;
  border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  mso-border-bottom-alt:solid windowtext .5pt;mso-border-right-alt:solid windowtext .5pt;
  padding:0cm 3.5pt 0cm 3.5pt;height:2.25pt'> <p class=MsoNormal><span style='font-size:8.0pt;font-family:Arial'><o:p><strong> 
            <%response.write (RS2("CO_Matricula_Escola"))%>
            - 
            <%response.write (RS2("NO_Aluno"))

Dim x
Dim strBarCode, strC
strBarCode = rsblo("CO_Barras")

If len(strBarCode) = 0 then strBarCode = "TEST"
%>
            </strong> <br>
            <%response.write (RS2("NO_Serie"))%>
            - 
            <%response.write (RS2("NO_Grau"))%>
            - Turma: 
            <%response.write (RS2("CO_Turma"))%>
            </o:p></span><span style='font-size:8.0pt;font-family:Arial'><o:p></o:p></span><span style='font-size:8.0pt;font-family:Arial'><o:p></o:p></span><span style='font-size:8.0pt;font-family:Arial'><o:p><strong> 
            </strong></o:p></span><span style='font-size:8.0pt;font-family:Arial'><o:p></o:p></span></p>
          </td>
      </tr>
      <td height="7" colspan=16 style='width:318.05pt;border:none;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:10pt'> </td>
      </tr>
      <tr style='mso-yfti-irow:34;height:18.5pt'> 
        <td colspan=10 valign=top style='width:318.05pt;border:none;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:18.5pt'> <p class=MsoNormal> 
            <%


WBarCode(rsblo("CO_Barras"))


'Rotina para gerar códigos de barra padrão 2of5 ou 25.

Sub WBarCode( Valor )
Dim f, f1, f2, i
Dim texto
Const fino = 1
Const largo = 3
Const altura = 50
Dim BarCodes(99)

if isempty(BarCodes(0)) then
  BarCodes(0) = "00110"
  BarCodes(1) = "10001"
  BarCodes(2) = "01001"
  BarCodes(3) = "11000"
  BarCodes(4) = "00101"
  BarCodes(5) = "10100"
  BarCodes(6) = "01100"
  BarCodes(7) = "00011"
  BarCodes(8) = "10010"
  BarCodes(9) = "01010"
  for f1 = 9 to 0 step -1
    for f2 = 9 to 0 Step -1
      f = f1 * 10 + f2
      texto = ""
      for i = 1 To 5
        texto = texto & mid(BarCodes(f1), i, 1) + mid(BarCodes(f2), i, 1)
      next
      BarCodes(f) = texto
    next
  next
end if

'Desenho da barra


' Guarda inicial
%>
            <img src=barcodes/p.gif width=<%=fino%> height=<%=altura%> border=0><img 
src=barcodes/b.gif width=<%=fino%> height=<%=altura%> border=0><img 
src=barcodes/p.gif width=<%=fino%> height=<%=altura%> border=0><img 
src=barcodes/b.gif width=<%=fino%> height=<%=altura%> border=0><img 

<%
texto = valor
if len( texto ) mod 2 <> 0 then
  texto = "0" & texto
end if


' Draw dos dados
do while len(texto) > 0
  i = cint( left( texto, 2) )
  texto = right( texto, len( texto ) - 2)
  f = BarCodes(i)
  for i = 1 to 10 step 2
    if mid(f, i, 1) = "0" then
      f1 = fino
    else
      f1 = largo
    end if
    %>
    src=barcodes/p.gif width=<%=f1%> height=<%=altura%> border=0><img 
    <%
    if mid(f, i + 1, 1) = "0" Then
      f2 = fino
    else
      f2 = largo
    end if
    %>
    src=barcodes/b.gif width=<%=f2%> height=<%=altura%> border=0><img 
    <%
  next
loop

' Draw guarda final
%>
src=barcodes/p.gif width=<%=largo%> height=<%=altura%> border=0><img 
src=barcodes/b.gif width=<%=fino%> height=<%=altura%> border=0><img 
src=barcodes/p.gif width=<%=1%> height=<%=altura%> border=0> 
            <%
end sub
%>
          </p></td>
        <td colspan=6 valign=top style='width:177.4pt;border:none;
  mso-border-top-alt:solid windowtext .5pt;padding:0cm 3.5pt 0cm 3.5pt;
  height:18.5pt'> <p class=MsoNormal><span style='font-size:7.0pt;font-family:Arial'>Autenticação 
            Mecânica – Ficha de Compensação</span><span style='font-size:7.0pt'><o:p></o:p></span></p></td>
      </tr>
      <tr valign="bottom" style='mso-yfti-irow:35;height:10.5pt'> 
        <td colspan=16 style='width:495.45pt;border:none;
  padding:0cm 3.5pt 0cm 3.5pt;height:10.5pt'> <p class=MsoNormal><span style='font-size:8.0pt;font-family:Arial'>........................................................................................................................................................................................................................</span><span
  style='font-size:7.0pt;font-family:Arial'><o:p></o:p></span><span style='font-size:8.0pt;font-family:Arial'>.....</span></p></td>
      </tr>
      <tr style='mso-yfti-irow:36;mso-yfti-lastrow:yes;height:10pt'> 
        <td colspan=16 valign=top style='width:495.45pt;border:none;
  padding:0cm 3.5pt 0cm 3.5pt;height:7.7pt'> <p class=MsoNormal><i><span style='font-size:7.5pt;font-family:Arial'>Recortar 
            na linha pontilhada abaixo do código de barras</span></i></p></td>
      </tr>
      <![if !supportMisalignedColumns]>
      <tr height=0> 
        <td width=117 style='border:none'></td>
        <td width=5 style='border:none'></td>
        <td width=53 style='border:none'></td>
        <td width=15 style='border:none'></td>
        <td width=48 style='border:none'></td>
        <td width=0 style='border:none'></td>
        <td width=98 style='border:none'></td>
        <td width=9 style='border:none'></td>
        <td width=46 style='border:none'></td>
        <td width=26 style='border:none'></td>
        <td width=39 style='border:none'></td>
        <td width=44 style='border:none'></td>
        <td width=14 style='border:none'></td>
        <td width=63 style='border:none'></td>
        <td width=38 style='border:none'></td>
        <td width=49 style='border:none'></td>
      </tr>
      <![endif]>
    </table>

</div>

<p class=MsoNormal><o:p>&nbsp;</o:p></p>

<p class=MsoNormal><o:p>&nbsp;</o:p></p>

<p class=MsoNormal><o:p>&nbsp;</o:p></p>

</div>

</body>

</html>
