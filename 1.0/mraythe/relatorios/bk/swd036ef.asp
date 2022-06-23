<%'On Error Resume Next%>
<%
Server.ScriptTimeout = 600 'valor em segundos
'Histórico escolar

%>
<!--#include file="../inc/caminhos.asp"-->
<!--#include file="../inc/parametros.asp"-->
<!--#include file="../inc/utils.asp"-->
<!--#include file="../inc/funcoes.asp"-->
<!--#include file="../inc/funcoes2.asp"-->
<!--#include file="../inc/funcoes6.asp"-->
<!--#include file="../inc/calculos.asp"-->
<!--#include file="../inc/resultados.asp"-->
<!--#include file="../../global/funcoes_diversas.asp"-->
<!--#include file="../inc/bd_historico.asp"-->

<% 
response.Charset="ISO-8859-1"
dados= request.QueryString("obr")
tipo= request.QueryString("tipo")
'if tipo = "EFA" then
'	tipo_cabec = 1
'	arquivo="SWD036FA"
'elseif tipo = "EFS" then
'	tipo_cabec = 2
'	arquivo="SWD036FS"	
'End if
tipo_cabec = 1
arquivo="SWD036EF"

nivel=4
permissao = session("permissao") 

sistema_local=session("sistema_local")
nvg=session("nvg")
session("nvg")=nvg
ano = DatePart("yyyy", now)
mes = DatePart("m", now) 
dia = DatePart("d", now) 
hora = DatePart("h", now) 
min = DatePart("n", now) 


vetor_historico = split(dados, "$!$")	

'if ori="ebe" then
'origem="../ws/doc/ofc/ebe/"
'end if

if mes<10 then
mes="0"&mes
end if

data = dia &"/"& mes &"/"& ano

if min<10 then
min="0"&min
end if

horario = hora & ":"& min


	'Dim AspPdf, Doc, Page, Font, Text, Param, Image, CharsPrinted
	'Instancia o objeto na memória
	SET Pdf = Server.CreateObject("Persits.Pdf")
	SET Doc = Pdf.CreateDocument
	Set Logo = Doc.OpenImage( Server.MapPath( "../img/logo_pdf.gif") )
	Set Font = Doc.Fonts.LoadFromFile(Server.MapPath("../fonts/arial.ttf"))	
	Set Font_Tesoura = Doc.Fonts.LoadFromFile(Server.MapPath("../fonts/ZapfDingbats.ttf"))
	If Font.Embedding = 2 Then
	   Response.Write "Embedding of this font is prohibited."
	   Set Font = Nothing
	End If
	If Font_Tesoura.Embedding = 2 Then
	   Response.Write "Embedding of this font is prohibited."
	   Set Font = Nothing
	End If 			 		 

		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR

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

		Set CONt = Server.CreateObject("ADODB.Connection") 
		ABRIRt = "DBQ="& CAMINHO_t & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONt.Open ABRIRt
    
	ano_letivo = vetor_historico(0)
	vetor_matriculas = vetor_historico(2)
	alunos_encontrados = split(vetor_matriculas, "#!#" )	
	

For alne=0 to ubound(alunos_encontrados)	
	cod_cons=alunos_encontrados(alne)

		
	Set RS0 = Server.CreateObject("ADODB.Recordset")
	SQL0 = "SELECT * FROM TB_Periodo ORDER BY NU_Periodo"
	RS0.Open SQL0, CON0
	check_periodo=1
	WHILE NOT RS0.EOF
		periodo=RS0("NU_Periodo")
		if check_periodo=1 then
			vetor_periodo=periodo
		else
			vetor_periodo=vetor_periodo&"#!#"&periodo
		end if
		check_periodo=check_periodo+1 
	RS0.MOVENEXT
	WEND				
		
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod_cons
	RS.Open SQL, CON1
	
	nome_aluno = RS("NO_Aluno")
	sexo_aluno = RS("IN_Sexo")
	nome_pai = RS("NO_Pai")
	nome_mae = RS("NO_Mae")
	co_uf_natural= RS("SG_UF_Natural")	
	co_cid_natural= RS("CO_Municipio_Natural")	
	inep = RS("RA_Aluno")	
	co_nacionalidade = RS("CO_Nacionalidade")	
	nome_aluno=replace_latin_char(nome_aluno,"html")	

	if nome_pai="" or isnull(nome_pai) then
		nome_pai=" "
	else
		nome_pai=replace_latin_char(nome_pai,"html")	
	end if
	
	if nome_mae="" or isnull(nome_mae) then
		nome_mae=" "
	else	
		nome_mae=replace_latin_char(nome_mae,"html")
	end if	
	
	if sexo_aluno="F" then
		desinencia="a"
	else
		desinencia="o"
	end if
				
	Set RS2 = Server.CreateObject("ADODB.Recordset")
	SQL2 = "SELECT * FROM TB_Contatos WHERE CO_Matricula="& cod_cons &" AND TP_Contato='ALUNO'"
	RS2.Open SQL2, CONCONT

	nascimento=RS2("DA_Nascimento_Contato")
	cod_cid_res=RS2("CO_Municipio_Res")
	cod_uf_res=RS2("SG_UF_Res")
	
	Set RS3n= Server.CreateObject("ADODB.Recordset")
	SQL3n = "SELECT * FROM TB_Municipios WHERE SG_UF='"& co_uf_natural &"' AND CO_Municipio="&co_cid_natural

	RS3n.Open SQL3n, CON0
	
	cid_nat=RS3n("NO_Municipio")		
	
	if not isnull(cid_nat) then
		cid_nat=replace_latin_char(cid_nat,"html")	
	end if				
	
	Set RS3 = Server.CreateObject("ADODB.Recordset")
	SQL3 = "SELECT * FROM TB_UF WHERE SG_UF='"& co_uf_natural &"'"
	RS3.Open SQL3, CON0
	
    if RS3.EOF then
	   uf_nat = "&nbsp;"
	else
	   uf_nat = RS3("NO_UF")				
	end if  

	if co_nacionalidade<>"" or not isnull(co_nacionalidade) then
	
			Set RS3 = Server.CreateObject("ADODB.Recordset")
			SQL3 = "SELECT * FROM TB_Nacionalidades WHERE CO_Nacionalidade="& co_nacionalidade 
			RS3.Open SQL3, CON0
		
			if RS3.EOF then
			   nacionalidade = "&nbsp;"
			else
			   nacionalidade = RS3("TX_Nacionalidade")				
			end if   
	end if  			

SET Page = Doc.Pages.Add( 595, 842 )


Set RS3 = Server.CreateObject("ADODB.Recordset")
SQL3 = "SELECT * FROM TB_Cabecalhos WHERE CO_Documento="& tipo_cabec
RS3.Open SQL3, CON7	



IF NOT RS3.EOF then
	cabec_1 = RS3("NO_Cabec_1")
	cabec_2 = RS3("NO_Cabec_2")
	cabec_3 = RS3("NO_Cabec_3")
	cabec_4 = RS3("NO_Cabec_4")
	cabec_5 = RS3("NO_Cabec_5")
	cabec_6 = RS3("NO_Cabec_6")
	cabec_7 = RS3("NO_Cabec_7")
	cabec_8 = RS3("NO_Cabec_8")	

				
	tit = "<i><font style=""font-size:12pt;""><b>"&cabec_1&"</b></font><br><font style=""font-size:8pt;"">	"&cabec_2&"<br>"&cabec_3&"<br>"&cabec_4&"<br>"&cabec_5&"<br>"&cabec_6&"<br>"&cabec_7&"<br>"&cabec_8&"<br></i>"																				
end if

'CABEÇALHO==========================================================================================		
Set Param_Logo_Gde = Pdf.CreateParam
margem=30				
largura_logo_gde=formatnumber(Logo.Width*0.5,0)
altura_logo_gde=formatnumber(Logo.Height*0.5,0)
area_utilizavel=Page.Width - (margem)*2	
Param_Logo_Gde("x") = margem
Param_Logo_Gde("y") = Page.Height - altura_logo_gde -22
Param_Logo_Gde("ScaleX") = 0.5
Param_Logo_Gde("ScaleY") = 0.5
Page.Canvas.DrawImage Logo, Param_Logo_Gde

'x_texto=largura_logo_gde+ 30
x_texto= margem+largura_logo_gde
y_texto=formatnumber(Page.Height - altura_logo_gde/3,0)
width_texto=Page.Width - 100

altura = altura_logo_gde+50
SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
Text = "<p>"&tit&"</p>"


Do While Len(Text) > 0
	CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
 
	If CharsPrinted = Len(Text) Then Exit Do
		SET Page = Page.NextPage
	Text = Right( Text, Len(Text) - CharsPrinted)
Loop 

x_texto= margem
y_texto=y_texto-90
width_texto=Page.Width

SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
Text = "<p><font style=""font-size:18pt;""><center>Hist&oacute;rico Escolar - Ensino Fundamental</center></font></p>"


Do While Len(Text) > 0
	CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
 
	If CharsPrinted = Len(Text) Then Exit Do
		SET Page = Page.NextPage
	Text = Right( Text, Len(Text) - CharsPrinted)
Loop 

'================================================================================================================			
Page.Canvas.SetParams "LineWidth=2" 
Page.Canvas.SetParams "LineCap=0" 
altura_primeiro_separador= y_texto -altura+65
With Page.Canvas
   .MoveTo margem, altura_primeiro_separador
   .LineTo Page.Width - margem, altura_primeiro_separador
   .Stroke
End With 

altura_segundo_separador= altura_primeiro_separador -60
With Page.Canvas
   .MoveTo margem, altura_segundo_separador
   .LineTo Page.Width - margem, altura_segundo_separador
   .Stroke
End With 			

Set param_table1 = Pdf.CreateParam("width=533; height=72; rows=6; cols=8; border=0; cellborder=0; cellspacing=0;")
Set Table = Doc.CreateTable(param_table1)
Table.Font = Font
y_primeira_tabela=altura_primeiro_separador-7
x_primeira_tabela=margem+5
With Table.Rows(1)
   .Cells(1).Width = 80
   .Cells(2).Width = 200  
   .Cells(3).Width = 20	
   .Cells(4).Width = 70	
   .Cells(5).Width = 20	
   .Cells(6).Width = 50	
   .Cells(7).Width = 40	
   .Cells(8).Width = 50				   		   		   		   
End With
Table(1, 2).ColSpan = 5
Table(2, 2).ColSpan = 5
Table(3, 2).ColSpan = 5
Table(4, 2).ColSpan = 5
Table(5, 3).ColSpan = 2	
Table(5, 5).ColSpan = 3													
Table(6, 2).ColSpan = 2
Table(6, 4).ColSpan = 3	


Table(1, 1).AddText "Alun"&desinencia&":", "size=8;", Font 
Table(1, 2).AddText "<b>"&nome_aluno&"</b>", "size=8;html=true", Font 
Table(1, 7).AddText "<div align=""right"">Matr&iacute;cula:</div>", "size=8;html=true", Font 
Table(1, 8).AddText "<div align=""right""><b>"&cod_cons&"</b></div>", "size=7;html=true", Font 	
Table(2, 1).AddText "Data de Nascimento:", "size=8;", Font 		
Table(2, 2).AddText formata(nascimento,"DD/MM/YYYY")&" &nbsp;&nbsp;&nbsp;&nbsp; Cidade: "&cid_nat&" &nbsp;&nbsp;&nbsp;&nbsp; Estado: "&uf_nat&" &nbsp;&nbsp;&nbsp;&nbsp; Nacionalidade: "&nacionalidade, "size=8;html=true", Font 	
Table(2, 7).AddText "<div align=""right"">INEP:</div>", "size=8;html=true", Font 
Table(2, 8).AddText "<div align=""right""><b>"&inep&"</b></div>", "size=7;html=true", Font
Table(3, 1).AddText "Pai: ", "size=8;", Font 
Table(3, 2).AddText nome_pai, "size=8;html=true", Font 
Table(4, 1).AddText "M&atilde;e: ", "size=8;html=true", Font 
Table(4, 2).AddText nome_mae, "size=8;html=true", Font	

Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 		


altura_medias=25
wrk_row = 2
wrk_altura_row = 12
altura_total = altura_medias			
Set param_table2 = Pdf.CreateParam("width=533; height="&altura_medias&"; rows="&wrk_row&"; cols=19; border=1; cellborder=0.1; cellspacing=0;")
Set Notas_Tit = Doc.CreateTable(param_table2)
Notas_Tit.Font = Font
y_medias=y_primeira_tabela-50-10

With Notas_Tit.Rows(1)
   .Cells(1).Width = 119
   .Cells(2).Width = 23
   .Cells(3).Width = 23
   .Cells(4).Width = 23
   .Cells(5).Width = 23
   .Cells(6).Width = 23
   .Cells(7).Width = 23
   .Cells(8).Width = 23   
   .Cells(9).Width = 23
   .Cells(10).Width = 23
   .Cells(11).Width = 23   
   .Cells(12).Width = 23 
   .Cells(13).Width = 23 
   .Cells(14).Width = 23 	
   .Cells(15).Width = 23 	
   .Cells(16).Width = 23 	
   .Cells(17).Width = 23 	
   .Cells(18).Width = 23
   .Cells(19).Width = 23			   
End With
Notas_Tit(1, 1).RowSpan = 2			

colSpan = 2																																				
Notas_Tit(1, 2).ColSpan = colSpan		
Notas_Tit(1, 4).ColSpan = colSpan	
Notas_Tit(1, 6).ColSpan = colSpan	
Notas_Tit(1, 8).ColSpan = colSpan	
Notas_Tit(1, 10).ColSpan = colSpan	
Notas_Tit(1, 12).ColSpan = colSpan	
Notas_Tit(1, 14).ColSpan = colSpan	
Notas_Tit(1, 16).ColSpan = colSpan		
Notas_Tit(1, 18).ColSpan = colSpan																																																								
Notas_Tit(1, 1).AddText "<div align=""center""><b>Disciplinas</b></div>", "size=10;indenty=4; html=true", Font 

wrkEtapaAnoHistorico = etapaAnoHistorico (tipo, "S")

vetorEtapaAnoHistorico = split(wrkEtapaAnoHistorico,"#!#")
numCol = 2
for e=0 to ubound(vetorEtapaAnoHistorico)
	Notas_Tit(1, numCol).AddText "<div align=""center"">"&vetorEtapaAnoHistorico(e)&"</div>", "size=6;alignment=center; indenty=2;html=true", Font 
	Notas_Tit(2, numCol).AddText "<div align=""center"">Conc.</div>", "size=6;alignment=center;indenty=2;html=true", Font 
	Notas_Tit(2, numCol+1).AddText "<div align=""center"">C.H.</div>", "size=6;alignment=center;indenty=2;html=true", Font 
	numCol=numCol+colSpan								
next

nomDisciplinas = HistoricoDisciplinas (cod_cons, tipo)
if isnull(nomDisciplinas) then
	qtd_disciplinas= 9		
else
	vetorDisciplinas = split(nomDisciplinas,"$!$")
	qtd_disciplinas= ubound(vetorDisciplinas)				
end if
anos = anoHistorico (cod_cons,tipo)
vetorAnos = split(anos,"#!#")
coluna=0	
wrk_row_inicial = wrk_row
wrk_loops=1
for anos = 0 to ubound(vetorAnos)
	seqAno = vetorAnos(anos)	

	wrk_row =wrk_row_inicial 
	
	wrkSomaCarga = 0	
	wrkSomaFrequencia = 0	
	coluna=coluna+2											
	for d = 0 to qtd_disciplinas
		wrk_loops=wrk_loops*1
		if wrk_loops>1 then
			wrk_row = wrk_row+1
		else
			Set Row = Notas_Tit.Rows.Add(wrk_altura_row)
			altura_total = altura_total+wrk_altura_row

			wrk_row = wrk_row+1
				
			Notas_Tit(wrk_row, 1).AddText "<div align=""left"">"&UCASE(vetorDisciplinas(d))&"</div>", "size=7;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font 	
		end if											
			dadosDisciplina = tbHistoricoAnoNota (cod_cons, seqAno, vetor_historico(1), vetorDisciplinas(d))
			vetorDadosDisciplina = split(dadosDisciplina,"#!#")
		
			qtd_dados = ubound(vetorDadosDisciplina)
			dados = ""
			if qtd_dados = -1 then
				wrkCargaHoraria = "&mdash;"	
				wrkFrequencia = ""
				wrkNota = "&mdash;"	
				inAprov = ""													
			else
				wrkCargaHoraria = vetorDadosDisciplina(0)	
				wrkFrequencia = vetorDadosDisciplina(1)	
				wrkNota = vetorDadosDisciplina(2)		
				inAprov = vetorDadosDisciplina(3)	
				if not isnull(wrkCargaHoraria) and wrkCargaHoraria<>"" and wrkCargaHoraria<>"X" then
					wrkSomaCarga = wrkSomaCarga+wrkCargaHoraria	
				end if	
				if not isnull(wrkFrequencia) and wrkFrequencia<>"" then							
					wrkSomaFrequencia = wrkSomaFrequencia+wrkFrequencia	
				end if								
				if inAprov="S" then
					wrkAprov= "APROVADO"
				else
					wrkAprov= "REPROVADO"			
				end if																						
			end if						
			Notas_Tit(wrk_row, coluna).AddText "<div align=""center"">"&wrkNota&"</div>", "size=7;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font 	
			coluna_disc=coluna+1
									
			Notas_Tit(wrk_row, coluna_disc).AddText "<div align=""center"">"&wrkCargaHoraria&"</div>", "size=7;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font			
	next
	if wrk_loops = 1 then
		wrkcargaAno = wrkSomaCarga
		wrkfreqAno = wrkSomaFrequencia
	else
		wrkcargaAno = wrkcargaAno&"#!#"&wrkSomaCarga
		wrkfreqAno = wrkfreqAno&"#!#"&wrkSomaFrequencia
	end if
	wrk_loops=wrk_loops+1						
	'seqAno = seqAno+1																		
next			
Set Row = Notas_Tit.Rows.Add(wrk_altura_row)
altura_total = altura_total+wrk_altura_row			
wrk_row = wrk_row+1		
Notas_Tit(wrk_row, 2).ColSpan = colSpan		
Notas_Tit(wrk_row, 4).ColSpan = colSpan	
Notas_Tit(wrk_row, 6).ColSpan = colSpan	
Notas_Tit(wrk_row, 8).ColSpan = colSpan	
Notas_Tit(wrk_row, 10).ColSpan = colSpan	
Notas_Tit(wrk_row, 12).ColSpan = colSpan	
Notas_Tit(wrk_row, 14).ColSpan = colSpan	
Notas_Tit(wrk_row, 16).ColSpan = colSpan		
Notas_Tit(wrk_row, 18).ColSpan = colSpan							
Notas_Tit(wrk_row, 1).AddText "CARGA HOR&Aacute;RIA TOTAL ", "size=8;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font 
cargaAno = totaisHistorico (cod_cons, tipo, "C")


vetorWrkcargaAno = split(wrkcargaAno,"#!#")
vetorCargaAno = split(cargaAno,"#!#")	
coluna=2


for h=0 to ubound(vetorCargaAno)	
	cargaExibe = 0			
	if vetorCargaAno(h)="" or isnull(vetorCargaAno(h)) then
		if vetorWrkcargaAno(h)=0 then
			cargaExibe = ""
		else
			cargaExibe = vetorWrkcargaAno(h)
		end if	
	else
		if vetorCargaAno(h)=0 then	
			if vetorWrkcargaAno(h)=0 then
				cargaExibe = ""
			else
				cargaExibe = vetorWrkcargaAno(h)
			end if						
		else
			cargaExibe = vetorCargaAno(h)						
		end if
	end if					
	Notas_Tit(wrk_row, coluna).AddText "<div align=""center"">"&cargaExibe&"</div>", "size=7;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font 							
	coluna = coluna+colSpan				  
next


Set Row = Notas_Tit.Rows.Add(wrk_altura_row)
altura_total = altura_total+wrk_altura_row			
wrk_row = wrk_row+1		
Notas_Tit(wrk_row, 2).ColSpan = colSpan		
Notas_Tit(wrk_row, 4).ColSpan = colSpan	
Notas_Tit(wrk_row, 6).ColSpan = colSpan	
Notas_Tit(wrk_row, 8).ColSpan = colSpan	
Notas_Tit(wrk_row, 10).ColSpan = colSpan	
Notas_Tit(wrk_row, 12).ColSpan = colSpan	
Notas_Tit(wrk_row, 14).ColSpan = colSpan	
Notas_Tit(wrk_row, 16).ColSpan = colSpan		
Notas_Tit(wrk_row, 18).ColSpan = colSpan							
Notas_Tit(wrk_row, 1).AddText "DIAS LETIVOS ", "size=8;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font 
cargaAno = totaisHistorico (cod_cons, tipo, "N")

vetorCargaAno = split(cargaAno,"#!#")	
coluna=2


for h=0 to ubound(vetorCargaAno)	
	cargaExibe = 0			
	cargaExibe = vetorCargaAno(h)								
	Notas_Tit(wrk_row, coluna).AddText "<div align=""center"">"&cargaExibe&"</div>", "size=7;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font 							
	coluna = coluna+colSpan				  
next







Set Row = Notas_Tit.Rows.Add(wrk_altura_row)
altura_total = altura_total+wrk_altura_row			
wrk_row = wrk_row+1
Notas_Tit(wrk_row, 2).ColSpan = colSpan		
Notas_Tit(wrk_row, 4).ColSpan = colSpan	
Notas_Tit(wrk_row, 6).ColSpan = colSpan	
Notas_Tit(wrk_row, 8).ColSpan = colSpan	
Notas_Tit(wrk_row, 10).ColSpan = colSpan	
Notas_Tit(wrk_row, 12).ColSpan = colSpan	
Notas_Tit(wrk_row, 14).ColSpan = colSpan	
Notas_Tit(wrk_row, 16).ColSpan = colSpan		
Notas_Tit(wrk_row, 18).ColSpan = colSpan				
Notas_Tit(wrk_row, 1).AddText "% FREQ&Uuml;&Ecirc;NCIA", "size=8;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font 			

freqAno = totaisHistorico (cod_cons, tipo, "F")

vetorWrkfreqAno = split(wrkfreqAno,"#!#")				
vetorFreqAno = split(freqAno,"#!#")	
coluna=2

for f=0 to ubound(vetorFreqAno)

	freqExibe = ""
	if vetorFreqAno(f)	="" or isnull(vetorFreqAno(f)) then
		if vetorWrkfreqAno(f)>0 then
			freqExibe = vetorWrkfreqAno(f)	
		end if		
	else
		if isnumeric(vetorFreqAno(f)) then					
			if vetorFreqAno(f) = 0 then	
				if vetorWrkfreqAno(f)>0 then
					freqExibe = vetorWrkfreqAno(f)	
				end if			
			else
				freqExibe = vetorFreqAno(f)						
			end if
		else
			freqExibe = vetorFreqAno(f)							
		end if
	end if				
	Notas_Tit(wrk_row, coluna).AddText "<div align=""center"">"&freqExibe&"</div>", "size=7;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font 
	coluna = coluna+colSpan				  							
next
Page.Canvas.DrawTable Notas_Tit, "x=30, y="&y_medias&"" 
y_tabela_2 = y_medias - altura_total - 20

if y_tabela_2< 200 then
	SET Page = Doc.Pages.Add( 595, 842 )
	
	'NOVO CABEÇALHO==========================================================================================		
	Set Param_Logo_Gde = Pdf.CreateParam
	margem=30				
	largura_logo_gde=formatnumber(Logo.Width*0.5,0)
	altura_logo_gde=formatnumber(Logo.Height*0.5,0)
	area_utilizavel=Page.Width - (margem)*2	
	Param_Logo_Gde("x") = margem
	Param_Logo_Gde("y") = Page.Height - altura_logo_gde -22
	Param_Logo_Gde("ScaleX") = 0.5
	Param_Logo_Gde("ScaleY") = 0.5
	Page.Canvas.DrawImage Logo, Param_Logo_Gde

	'x_texto=largura_logo_gde+ 30
	x_texto= margem+largura_logo_gde
	y_texto=formatnumber(Page.Height - altura_logo_gde/3,0)
	width_texto=Page.Width - 100

	altura = altura_logo_gde+50
	SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
	Text = "<p>"&tit&"</p>"


	Do While Len(Text) > 0
		CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
	 
		If CharsPrinted = Len(Text) Then Exit Do
			SET Page = Page.NextPage
		Text = Right( Text, Len(Text) - CharsPrinted)
	Loop 

	x_texto= margem
	y_texto=y_texto-90
	width_texto=Page.Width

	SET Param = Pdf.CreateParam("x="&x_texto&";y="&y_texto&"; height="&altura&"; width="&width_texto&"; alignment=center; size=14; color=#000000; html=true")
	Text = "<p><font style=""font-size:18pt;""><center>Hist&oacute;rico Escolar - Ensino Fundamental</center></font></p>"


	Do While Len(Text) > 0
		CharsPrinted = Page.Canvas.DrawText(Text, Param, Font )
	 
		If CharsPrinted = Len(Text) Then Exit Do
			SET Page = Page.NextPage
		Text = Right( Text, Len(Text) - CharsPrinted)
	Loop 

	'================================================================================================================		
	Page.Canvas.SetParams "LineWidth=2" 
	Page.Canvas.SetParams "LineCap=0" 
	altura_primeiro_separador= y_texto -altura+65
	With Page.Canvas
	   .MoveTo margem, altura_primeiro_separador
	   .LineTo Page.Width - margem, altura_primeiro_separador
	   .Stroke
	End With 

	altura_segundo_separador= altura_primeiro_separador -60
	With Page.Canvas
	   .MoveTo margem, altura_segundo_separador
	   .LineTo Page.Width - margem, altura_segundo_separador
	   .Stroke
	End With 			

	Set param_table1 = Pdf.CreateParam("width=533; height=72; rows=6; cols=8; border=0; cellborder=0; cellspacing=0;")
	Set Table = Doc.CreateTable(param_table1)
	Table.Font = Font
	y_primeira_tabela=altura_primeiro_separador-7
	x_primeira_tabela=margem+5
	With Table.Rows(1)
	   .Cells(1).Width = 80
	   .Cells(2).Width = 200  
	   .Cells(3).Width = 20	
	   .Cells(4).Width = 70	
	   .Cells(5).Width = 20	
	   .Cells(6).Width = 50	
	   .Cells(7).Width = 40	
	   .Cells(8).Width = 50				   		   		   		   
	End With
	Table(1, 2).ColSpan = 5
	Table(2, 2).ColSpan = 5
	Table(3, 2).ColSpan = 5
	Table(4, 2).ColSpan = 5
	Table(5, 3).ColSpan = 2	
	Table(5, 5).ColSpan = 3													
	Table(6, 2).ColSpan = 2
	Table(6, 4).ColSpan = 3	


	Table(1, 1).AddText "Alun"&desinencia&":", "size=8;", Font 
	Table(1, 2).AddText "<b>"&nome_aluno&"</b>", "size=8;html=true", Font 
	Table(1, 7).AddText "<div align=""right"">Matr&iacute;cula:</div>", "size=8;html=true", Font 
	Table(1, 8).AddText "<div align=""right""><b>"&cod_cons&"</b></div>", "size=8;html=true", Font 	
	Table(2, 1).AddText "Data de Nascimento:", "size=8;", Font 		
	Table(2, 2).AddText formata(nascimento,"DD/MM/YYYY")&" &nbsp;&nbsp;&nbsp;&nbsp; Cidade: "&cid_nat&" &nbsp;&nbsp;&nbsp;&nbsp; Estado: "&uf_nat&" &nbsp;&nbsp;&nbsp;&nbsp; Nacionalidade: "&nacionalidade, "size=8;html=true", Font 	
	Table(2, 7).AddText "<div align=""right"">INEP:</div>", "size=8;html=true", Font 
	Table(2, 8).AddText "<div align=""right""><b>"&inep&"</b></div>", "size=8;html=true", Font							
	Table(3, 1).AddText "Pai: ", "size=8;", Font 
	Table(3, 2).AddText nome_pai, "size=8;html=true", Font 
	Table(4, 1).AddText "M&atilde;e: ", "size=8;html=true", Font 
	Table(4, 2).AddText nome_mae, "size=8;html=true", Font	

	Page.Canvas.DrawTable Table, "x="&x_primeira_tabela&", y="&y_primeira_tabela&"" 		


	altura_medias=25
	wrk_row = 2
	wrk_altura_row = 12
	altura_total = altura_medias			
	Set param_table2 = Pdf.CreateParam("width=533; height="&altura_medias&"; rows="&wrk_row&"; cols=19; border=1; cellborder=0.1; cellspacing=0;")
	Set Notas_Tit = Doc.CreateTable(param_table2)
	Notas_Tit.Font = Font
	y_tabela_2=y_primeira_tabela-50-10	
end if
	
wrk_row =1
Set param_table2 = Pdf.CreateParam("width=533; height=15; rows="&wrk_row&"; cols=6; border=1; cellborder=0.1; cellspacing=0;")
Set Notas_Tit = Doc.CreateTable(param_table2)
Notas_Tit.Font = Font
y_medias=y_primeira_tabela-72-10	

With Notas_Tit.Rows(1)
   .Cells(1).Width = 50
   .Cells(2).Width = 50
   .Cells(3).Width = 170
   .Cells(4).Width = 83
   .Cells(5).Width = 30
   .Cells(6).Width = 150		   
End With
Notas_Tit(1, 1).AddText "<div align=""center""><B>Ano Letivo</B></div>", "size=7;indenty=2; html=true", Font 
if tipo = "EFA" then
	Notas_Tit(1, 2).AddText "<div align=""center""><B>Ano</B></div>", "size=7;indenty=2; html=true", Font 	
else
	Notas_Tit(1, 2).AddText "<div align=""center""><B>S&eacute;rie</B></div>", "size=7;indenty=2; html=true", Font 				
end if	
Notas_Tit(1, 3).AddText "<div align=""center""><B>Estabelecimento</B></div>", "size=7;indenty=2; html=true", Font 
Notas_Tit(1, 4).AddText "<div align=""center""><B>Munic&iacute;pio</B></div>", "size=7;indenty=2; html=true", Font 
Notas_Tit(1, 5).AddText "<div align=""center""><B>UF</B></div>", "size=7;indenty=2; html=true", Font 
Notas_Tit(1, 6).AddText "<div align=""center""><B>Resultado</B></div>", "size=7;indenty=2; html=true", Font 														

AnoEscola = tbHistoricoAnoEscola (cod_cons, tipo, NULL)

if isnull(AnoEscola) then
	anos= 9		
else
	vetorAnoEscola	= split(AnoEscola,"$!$")	
	anos= ubound(vetorAnoEscola)
end if

for a = 0 to anos
	Set Row = Notas_Tit.Rows.Add(10)
	wrk_row = wrk_row+1
	vetorAno = split(vetorAnoEscola(a),"#!#")	

	serie = etapaHistorico (tipo,vetorAno(1),"A")

	
	Notas_Tit(wrk_row, 1).AddText "<div align=""center"">"&vetorAno(0)&"</div>", "size=6;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font 			
	Notas_Tit(wrk_row, 2).AddText "<div align=""center"">"&serie&"</div>", "size=6;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font 			
	Notas_Tit(wrk_row, 3).AddText vetorAno(2), "size=6;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font 			
	Notas_Tit(wrk_row, 4).AddText "<div align=""center"">"&vetorAno(3)&"</div>", "size=6;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font 			
	Notas_Tit(wrk_row, 5).AddText "<div align=""center"">"&vetorAno(4)&"</div>", "size=6;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font 			
	Notas_Tit(wrk_row, 6).AddText "<div align=""center"">"&vetorAno(5)&"</div>", "size=6;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font 				

next		


wrkObsHistorico = observacaoHistorico (cod_cons, tipo)
if not (wrkObsHistorico="" or isnull(wrkObsHistorico)) then
	Set param_materias = PDF.CreateParam	
	param_materias.Set "expand=true" 
	Set Row = Notas_Tit.Rows.Add(15)
	wrk_row = wrk_row+1
	Notas_Tit(wrk_row, 1).ColSpan = 6	
	Notas_Tit(wrk_row, 1).AddText "<B>Observa&ccedil;&otilde;es:</B> "& wrkObsHistorico, "size=7;alignment=center;indentx=2;indenty=2;html=true;expand=true", Font 			
end if	
Page.Canvas.DrawTable Notas_Tit, "x=30, y=" &y_tabela_2				



	y_linha=50
	With Page.Canvas
	   .MoveTo 230, y_linha
	   .LineTo 370, y_linha
	   .Stroke
	End With 	

	With Page.Canvas
	   .MoveTo 403, y_linha
	   .LineTo 553, y_linha
	   .Stroke
	End With 
	
y_data = y_linha+20
'Data===========================================================================
	SET Param_data = Pdf.CreateParam("x="&margem&";y="&y_data&"; height=40; width=300; alignment=Left; size=8; color=#000000;html=true")
	data_formatada = FormatDateTime(Now(),1)
	data_formatada = Server.URLEncode(data_formatada)
	data_formatada = replace(data_formatada,"segunda%2Dfeira%2C","")
	data_formatada = replace(data_formatada,"ter%E7a%2Dfeira%2C","")
	data_formatada = replace(data_formatada,"quarta%2Dfeira%2C","")
	data_formatada = replace(data_formatada,"quinta%2Dfeira%2C","")
	data_formatada = replace(data_formatada,"sexta%2Dfeira%2C","")
	data_formatada = replace(data_formatada,"s%E1bado%2C","")
	data_formatada = replace(data_formatada,"domingo%2C","")
	data_formatada = replace(data_formatada,"+"," ")
	data_formatada = replace(data_formatada,"%E7", "&ccedil;")			
																			
data_preenche = "Rio de Janeiro, RJ, "&data_formatada 
CharsPrinted = Page.Canvas.DrawText(data_preenche, Param_data, Font )


'===========================================================================
Set RS3 = Server.CreateObject("ADODB.Recordset")
SQL3 = "SELECT * FROM TB_Assinatura"
RS3.Open SQL3, CON7		

IF RS3.EOF then
	secretaria = "Secret&aacute;rio(a)"
	diretor = "Diretor(a)"			
else
	secretaria =RS3("NO_Secretaria")&"<br>Secret&aacute;rio(a)<br>"&RS3("NU_Sec_Sec") 
	diretor = RS3("NO_Diretor")&"<br>Diretor(a)<br>"&RS3("NU_Sec_Dir")

end if		

 SET Param_secretario = Pdf.CreateParam("x=210;y="&y_linha&"; height=40; width=160; alignment=center; size=9; color=#000000;html=true")
secretario = "<div align=""center"">"&secretaria&"</div>"
CharsPrinted = Page.Canvas.DrawText(secretario, Param_secretario, Font )

 SET Param_diretor = Pdf.CreateParam("x=393;y="&y_linha&"; height=40; width=160; alignment=center; size=9; color=#000000;html=true")
diretor = "<div align=""center"">"&diretor&"</div>"
CharsPrinted = Page.Canvas.DrawText(diretor, Param_diretor, Font )
	
Next						

Doc.SaveHttp("attachment; filename=" & arquivo & ".pdf")   
%>
