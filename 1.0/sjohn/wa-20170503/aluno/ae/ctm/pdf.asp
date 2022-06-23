<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 30 'valor em segundos
%>
<!--#include file="../../../../inc/graficos.asp"-->
<% 
arquivo="Grafico"
response.Charset="ISO-8859-1"
opt= request.QueryString("opt")
ori= request.QueryString("ori")

nivel=4
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
nvg=session("nvg")
session("nvg")=nvg
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 

if ori="edf" then
origem="../ws/mat/man/eco/"
end if




if mes<10 then
mes="0"&mes
end if

data = dia &"/"& mes &"/"& ano

if mes=1 then
	mes_extenso="Janeiro"
elseif mes=2 then
	mes_extenso="Fevereiro"
elseif mes=3 then
	mes_extenso="Mar&ccedil;o"
elseif mes=4 then
	mes_extenso="Abril"
elseif mes=5 then
	mes_extenso="Maio"
elseif mes=6 then
	mes_extenso="Junho"
elseif mes=7 then
	mes_extenso="Julho"
elseif mes=8 then
	mes_extenso="Agosto"
elseif mes=9 then
	mes_extenso="Setembro"
elseif mes=10 then
	mes_extenso="Outubro"
elseif mes=11 then
	mes_extenso="Novembro"
elseif mes=12 then
	mes_extenso="Dezembro"
end if	
data_extenso="Rio de Janeiro, "&dia &" de "& mes_extenso &" de "& ano
if min<10 then
min="0"&min
end if

horario = hora & ":"& min

	'Dim AspPdf, Doc, Page, Font, Text, Param, Image, CharsPrinted
	'Instancia o objeto na mem&oacute;ria
	SET Pdf = Server.CreateObject("Persits.Pdf")
	SET Doc = Pdf.CreateDocument
	Set Logo = Doc.OpenImage( Server.MapPath( "../../../../img/logo_pdf.gif") )
	Set Font = Doc.Fonts.LoadFromFile(Server.MapPath("../../../../fonts/arial.ttf"))	
	Set Font_Tesoura = Doc.Fonts.LoadFromFile(Server.MapPath("../../../../fonts/ZapfDingbats.ttf"))
	If Font.Embedding = 2 Then
	   Response.Write "Embedding of this font is prohibited."
	   Set Font = Nothing
	End If
	If Font_Tesoura.Embedding = 2 Then
	   Response.Write "Embedding of this font is prohibited."
	   Set Font = Nothing
	End If 			 		 

		
	
SET Page = Doc.Pages.Add(842, 595)

'CABE&Ccedil;ALHO==========================================================================================		
		Set Param_Logo_Gde = Pdf.CreateParam

		largura_logo_gde=formatnumber(Logo.Width*0.3,0)
		altura_logo_gde=formatnumber(Logo.Height*0.3,0)
		margem=30			
		linha=10
		area_utilizavel=Page.Width-(margem*2)
		Param_Logo_Gde("x") = formatnumber(area_utilizavel/2,0)
		Param_Logo_Gde("y") = Page.Height - altura_logo_gde -margem
		Param_Logo_Gde("ScaleX") = 0.3
		Param_Logo_Gde("ScaleY") = 0.3
		Page.Canvas.DrawImage Logo, Param_Logo_Gde

		x_texto=largura_logo_gde+ margem+10
		y_posicao=595-margem 
		y_posicao=y_posicao - altura_logo_gde-linha
		
		SET Param = Pdf.CreateParam("x="&margem&";y="&y_posicao&"; height=20; width="&area_utilizavel&"; alignment=center; size=14; color=#000000; html=true")
		
		Text = "<center><b><font style=""font-size:15pt;"">GRÁFICO</FONT></b></center>" 			
		
		Do While Len(Text) > 0
			CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
		 
			If CharsPrinted = Len(Text) Then Exit Do
				SET Page = Page.NextPage
			Text = Right( Text, Len(Text) - CharsPrinted)
		Loop 
						

'FIM DO CABE&Ccedil;ALHO==========================================================================================	
		
	faixas=session("faixas")
categorias=session("categorias")
tp_grafico=session("tp_grafico")
img= ColunaAgrupada_2D_ou_3D(faixas,categorias,tp_grafico)	
'Set grafico = Doc.OpenImage(img)	

'		Param_grafico("x") = formatnumber(area_utilizavel/2,0)
'		Param_grafico("y") = Page.Height - altura_logo_gde -(margem*3)
'		Page.Canvas.DrawImage grafico, Param_grafico
			
	'Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
Filename = Doc.Save( Server.MapPath("grafico.pdf"), False )	

%>