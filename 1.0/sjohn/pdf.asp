<%Dim AspPdf, Doc, Page, Font, Text, Param, Filename, CharsPrinted
 
'Instancia o objeto na memória
SET AspPdf = Server.CreateObject("Persits.Pdf")
SET Doc = AspPdf.CreateDocument
 
'Define o tamanho da folha em milímetros
SET Page = Doc.Pages.Add( 210, 297 )
 
'Define o tipo de fonte a ser utilizada
SET Font = Doc.Fonts("Times-Roman")

'Define os parâmetros de alinhamento: X, Y do canto superior esquerdo ao lado inferior direito, altura, largura e tamanho da fonte.
SET param = AspPdf.CreateParam("x=10;y=270;height=260;width=196; size=10;")
'param.Add("x=10; y=20")
'param.Add("z=30")
'The method Set clears the object before initializing it to the new values. Therefore, the code 
'param.Set("x=10; y=20") 
'is equivalent to
'param.Clear
'param.Add("x=10; y=20") 


 
'Obtem o texto informado no formulário html
Text = "Teste"

Do While Len(Text) > 0
    CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
 
    If CharsPrinted = Len(Text) Then Exit Do
        SET Page = Page.NextPage
	Text = Right( Text, Len(Text) - CharsPrinted)
Loop 

 
'Define os parâmetros de alinhamento: X, Y do canto superior esquerdo ao lado inferior direito, altura, largura e tamanho da fonte.
SET param = AspPdf.CreateParam("x=10;y=270;height=260;width=196; size=10;")
'This line saves the document to disk. The first argument is required and must be set to a full file path. The second argument is optional. If set to True or omitted, it instructs the Save method to overwrite an existing file. If set to False, it forces unique file name generation. For example, if the file hello.pdf already exists in the specified directory and another document is being saved under the same name, the Save method tries the filenames hello(1).pdf, hello(2).pdf, etc., until a non-existent name is found. The Save method returns the filename (without the path) under which the file was saved. 
'Filename = Doc.Save( "e:\home\SEU_LOGIN_FTP\web\asppdf\hello.pdf", False )
'Para salvar em uma base de dados
'rs("FileBlob").Value = Doc.SaveToMemory
'Response.ContentType = "application/pdf" 

'** nome que o arquivo terá ao ser baixado, neste caso, tiramos o código.
arquivo="texto.pdf"
'Response.AddHeader "Content-Disposition", "attachment;filename=""" & arquivo & """" 
Filename = Doc.SaveHttp( Server.MapPath(arquivo), False )
'Response.BinaryWrite Filename
%>