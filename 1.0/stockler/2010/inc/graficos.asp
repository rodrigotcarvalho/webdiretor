<%
function Pizza(faixas,categorias)
response.Charset="ISO-8859-1"
Set oChart = CreateObject("OWC10.ChartSpace")
Set c = oChart.Constants
oChart.Border.Color = c.chColorNone
'Dim oChartSpace As OWC11.ChartSpaceClass = New OWC11.ChartSpaceClass
'Dim chartType As OWC11.ChartChartTypeEnum

valores=split(faixas,"#!#")
classes=split(categorias,"#!#")

Dim categories(20), Vals(20), Vals2(3)
for i=0 to ubound(valores)
' críe um array que represente os valores da série.
Vals(i) = valores(i)
next
' gráfico de pizza com quatro categorias.
' críe um array que represente as categorias.
for y=0 to ubound(classes)
categories(y) = classes(y)
next
With oChart
' adicionando um objeto do gráfico.
.Charts.Add
' adicionando o tipo do gráfico.
.Charts(0).Type = oChart.Constants.chChartTypePie3d
' adicionando a série ao gráfico.
.Charts(0).SeriesCollection.Add
' ajustando o subtítulo da série (o texto da legenda).
.Charts(0).SeriesCollection(0).Caption = ""
.Charts(0).SeriesCollection(0).TipText = ""
.Charts(0).SeriesCollection(0).DataLabelsCollection.Add()


' adicionando as categorias e os valoresda série.
.Charts(0).SeriesCollection(0).SetData c.chDimCategories, c.chDataLiteral, categories
.Charts(0).SeriesCollection(0).SetData c.chDimValues, c.chDataLiteral, Vals
.Charts(0).SeriesCollection(0).DataLabelsCollection(0).HasPercentage = FALSE
.Charts(0).SeriesCollection(0).DataLabelsCollection(0).HasValue = TRUE
.Charts(0).SeriesCollection(0).DataLabelsCollection(0).Font.Name = "verdana"
.Charts(0).SeriesCollection(0).DataLabelsCollection(0).Font.Size = 10
.Charts(0).SeriesCollection(0).DataLabelsCollection(0).Font.Bold = True
.Charts(0).SeriesCollection(0).DataLabelsCollection(0).Font.Color = "#666666"
.Charts(0).SeriesCollection(0).DataLabelsCollection(0).Position = c.chLabelPositionCenter

.Charts(0).HasLegend = True
.Charts(0).HasTitle = False
End With


Response.Expires = 0
Response.Buffer = true
Response.Clear
Response.ContentType = "image/gif"
'ajustando o tamanho do gráfico (figura).
Response.BinaryWrite oChart.GetPicture("gif",980, 300)
end function

'===========================================================================================================
function ColunaAgrupada(faixas,categorias)
response.Charset="ISO-8859-1"
Set oChart = CreateObject("OWC10.ChartSpace")
Set c = oChart.Constants
oChart.Border.Color = c.chColorNone



 
Dim categories(30), Vals(31), nome_turma(10)

'response.Write(faixas)

valores=split(faixas,"#$#")
classes=split(categorias,"#!#")
'response.Write(ubound(classes)&"<BR>")
conta_medias=0

		for y=0 to ubound(classes)
		
			if classes(y)="MED" then
				if conta_medias=0 then		
					categories(y) = classes(y)
					conta_medias=conta_medias+1
				elseif conta_medias=1 then
					categories(y) = "MED "
					conta_medias=conta_medias+1
				elseif conta_medias=2 then
					categories(y) = " MED "
					conta_medias=conta_medias+1	
				elseif conta_medias=4 then
					categories(y) = " MED  "
					conta_medias=conta_medias+1	
				elseif conta_medias=5 then
					categories(y) = "  MED  "
					conta_medias=conta_medias+1						
				end if	
			else
				categories(y) = classes(y)
			end if	
'response.Write(categories(y)&"==<BR>")		
		next	

With oChart
 
'adicionando um objeto do gráfico.
.Charts.Add 
 
'adicionando o tipo do gráfico.
.Charts(0).Type = oChart.Constants.chChartTypeColumnClustered3D
 
'response.Write(ubound(valores)&"<BR>")
'for n=0 to ubound(valores)
'response.Write(valores(n)&"==<BR>")
'next
'RESPONSE.END()
quantidade_de_turmas=ubound(valores)
 For k=0 to quantidade_de_turmas
	colunas=Split(valores(k),"#!#")
		nome_turma(k)=colunas(0)
	'response.Write(ubound(colunas))
	For m=1 to ubound(colunas)	
		Vals(m)=colunas(m)
		if Vals(m)="" or isnull(Vals(m)) then
			Vals(m)=0
		end if
	next
	
		'adicionando a primeira série ao gráfico.
		.Charts(0).SeriesCollection.Add 
		 
		'ajustando o subtítulo da série (o texto da legenda).
		.Charts(0).SeriesCollection(k).Caption = nome_turma(k)
		 
		'adicionando as categorias e os valores da primeira série.
		.Charts(0).SeriesCollection(k).SetData c.chDimCategories, c.chDataLiteral, categories
		.Charts(0).SeriesCollection(k).SetData c.chDimValues, c.chDataLiteral, Vals
	next 
End With

With oChart
 
.Charts(0).HasLegend = True
.Charts(0).Legend.Position = 4 '1 superior 2 inferior                
'.Charts(0).Legend.Font.Name = "tahoma"          
'.Charts(0).Legend.Font.Size = 8                
'.Charts(0).Legend.Border.Color = c.chColorNone 'borda

.Charts(0).HasTitle = False
'.Charts(0).Title.Font.Name = "Arial"
'.Charts(0).Title.Font.Size = 10
'.Charts(0).Title.Font.Bold = true
'.Charts(0).Axes(c.chAxisPositionValue).MajorGridlines.Line.Color = "Black"
'.Charts(0).Axes(c.chAxisPositionValue).MinorGridlines.Line.Color = "Gray"
'.Charts(0).Title.Caption = ucase(mensagem)
'.Charts(0).PlotArea.Interior.Color = "#ffffff"  
'.Charts(0).Type = 0
'.ChartLayout = c.chChartLayoutHorizontal
.Charts(0).Axes(0).Font.Size = 8
.Charts(0).Axes(0).Font.Name = "Tahoma"
'.Charts(0).Axes(0).Scaling.Minimum = 0
'.Charts(0).Axes(0).Scaling.Maximum = 100
'.Charts(0).Axes(0).HasTickLabels = True
'.Charts(0).Axes(1).HasTitle = true
'.Charts(0).Axes(1).Title.Font.Size = 8
'.Charts(0).Axes(1).Title.Font.Name = "Tahoma"
'.Charts(0).Axes(1).Title.Caption = "Notas"
.Charts(0).Axes(1).Font.Size = 8
.Charts(0).Axes(1).Font.Name = "Tahoma"
.Charts(0).Axes(1).Scaling.Minimum = 0
.Charts(0).Axes(1).Scaling.Maximum = 10
.Charts(0).Axes(1).HasTickLabels = True
End With


'axScale = oChart.Charts(0).Axes(ChartAxisPositionEnum.chAxisPositionValue).Scaling  
'axScale.Maximum = 100  
'axScale.Minimum = 0  
'axValAxis = oChart.Charts(0).Axes(ChartAxisPositionEnum.chAxisPositionValue)  
'axValAxis.HasMajorGridlines = False  
'axValAxis.HasMinorGridlines = False  

 
Response.Expires = 0
Response.Buffer = true
Response.Clear
Response.ContentType = "image/gif"
 
'ajustando o tamanho do gráfico (figura).
Response.BinaryWrite oChart.GetPicture("gif",980, 320) 
 
 
'Set objPieChart = Nothing
end function

'===========================================================================================================
function Coluna(faixas,categorias)
response.Charset="ISO-8859-1"
Set oChart = CreateObject("OWC10.ChartSpace")
Set c = oChart.Constants
oChart.Border.Color = c.chColorNone
 
Dim categories(30), Vals(31), nome_turma(10)

'response.Write(faixas)

valores=split(faixas,"#$#")
classes=split(categorias,"#!#")
'response.Write(ubound(classes)&"<BR>")

		for y=0 to ubound(classes)
			categories(y) = classes(y)
		next	

With oChart
 
'adicionando um objeto do gráfico.
.Charts.Add 
 
'adicionando o tipo do gráfico.
.Charts(0).Type = oChart.Constants.chChartTypeColumnClustered3D
 
'response.Write(ubound(valores)&"<BR>")
'for n=0 to ubound(valores)
'response.Write(valores(n)&"==<BR>")
'next
quantidade_de_turmas=ubound(valores)
'm=0
 For k=0 to quantidade_de_turmas
	colunas=Split(valores(k),"#!#")
		nome_turma(k)=colunas(0)
	'response.Write(ubound(colunas))
	for m=1 to ubound(colunas)
			Vals(m)=colunas(m)
'response.Write(Vals(m))		
	next		

	
		'adicionando a primeira série ao gráfico.
		.Charts(0).SeriesCollection.Add 
		 
		'ajustando o subtítulo da série (o texto da legenda).
		.Charts(0).SeriesCollection(k).Caption = nome_turma(k)	 
		'adicionando as categorias e os valores da primeira série.
		.Charts(0).SeriesCollection(k).SetData c.chDimCategories, c.chDataLiteral, categories
		.Charts(0).SeriesCollection(k).SetData c.chDimValues, c.chDataLiteral, Vals
	next 
End With
 
With oChart
 
.Charts(0).HasLegend = True
.Charts(0).HasTitle = False
 
End With
 
Response.Expires = 0
Response.Buffer = true
Response.Clear
Response.ContentType = "image/gif"
' 
'ajustando o tamanho do gráfico (figura).
Response.BinaryWrite oChart.GetPicture("gif",980, 320) 
 
 
'Set objPieChart = Nothing
end function

'===========================================================================================================
function StackedColuna(faixas,categorias)
response.Charset="ISO-8859-1"
Set oChart = CreateObject("OWC10.ChartSpace")
Set c = oChart.Constants
oChart.Border.Color = c.chColorNone
 
Dim categories(30), Vals(31), nome_turma(10)

'response.Write(faixas)

valores=split(faixas,"#$#")
classes=split(categorias,"#!#")
'response.Write(ubound(classes)&"<BR>")

		for y=0 to ubound(classes)
			categories(y) = classes(y)
		next	

With oChart
 
'adicionando um objeto do gráfico.
.Charts.Add 
 
'adicionando o tipo do gráfico.
.Charts(0).Type = oChart.Constants.chChartTypeColumnStacked3D
 
'response.Write(ubound(valores)&"<BR>")
'for n=0 to ubound(valores)
'response.Write(valores(n)&"==<BR>")
'next
quantidade_de_disciplinas=ubound(valores)
maximo_escala=2
 For k=0 to quantidade_de_disciplinas
	colunas=Split(valores(k),"#!#")
	nome_turma(k)=colunas(0)
	'response.Write(ubound(colunas))
	for m=1 to ubound(colunas)
		Vals(m)=colunas(m)
	next		
	
		Vals(1)=Vals(1)*1
		maximo_escala=maximo_escala*1
		maximo_escala=maximo_escala+Vals(1)

'response.Write(maximo_escala&"<br>")				
	
		'adicionando a primeira série ao gráfico.
		.Charts(0).SeriesCollection.Add 
		 
		'ajustando o subtítulo da série (o texto da legenda).
		.Charts(0).SeriesCollection(k).Caption = nome_turma(k)	 
		'adicionando as categorias e os valores da primeira série.
		.Charts(0).SeriesCollection(k).SetData c.chDimCategories, c.chDataLiteral, categories
		.Charts(0).SeriesCollection(k).SetData c.chDimValues, c.chDataLiteral, Vals
	next 
End With
 
With oChart
.Charts(0).Axes(1).Scaling.Minimum = 0
.Charts(0).Axes(1).Scaling.Maximum = maximo_escala
.Charts(0).HasLegend = True
.Charts(0).Legend.Position = 2 '1 superior 2 inferior 4 Direita   
.Charts(0).HasTitle = False
 
End With
 
Response.Expires = 0
Response.Buffer = true
Response.Clear
Response.ContentType = "image/gif"
' 
'ajustando o tamanho do gráfico (figura).
Response.BinaryWrite oChart.GetPicture("gif",980, 320) 
 
 
'Set objPieChart = Nothing
end function
%>

