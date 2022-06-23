<%
FUNCTION calculamedia(codigo,unidade,curso,co_etapa,turma,materia,periodo)
' MCAL

contanotas=0
verifica = "ok"

ACU_Peso1 = 0
ACU_Peso2 = 0
ACU_Peso3 = 0
ACU_Peso4 = 0
nota_princ1=0
nota_princ2=0
nota_princ3=0
nota_princ4=0			  
nu_peso2 = 1

ACU_Rec1=0
ACU_Rec2=0
ACU_Rec3=0
ACU_Rec4=0

acu_r1=0
acu_r2=0
acu_r3=0
acu_r4=0

mrec1=0
mrec2=0
mrec3=0
mrec4=0




'response.Write("SQLo = SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' And CO_Materia ='"& materia&"'" )
		
		Set RSo = Server.CreateObject("ADODB.Recordset")
		SQLo = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' And CO_Materia ='"& materia&"'" 
		RSo.Open SQLo, CON0
		
	mae=RSo("IN_MAE")
	fil=RSo("IN_FIL")
	in_co=RSo("IN_CO")
	nu_peso=RSo("NU_Peso")
	ordem=RSo("NU_Ordem_Boletim")
		
if mae=TRUE AND fil=false AND in_co=true AND isnull(nu_peso) then
	ordem2=ordem+1
	While verifica = "ok"

	'response.Write(">>"&ordem2)
	
		Set RSof = Server.CreateObject("ADODB.Recordset")
		SQLof = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' And NU_Ordem_Boletim="&ordem2 
		RSof.Open SQLof, CON0
		
		materia_fil=RSof("CO_Materia")
		mae=RSof("IN_MAE")
		fil=RSof("IN_FIL")
		in_co=RSof("IN_CO")
		nu_peso=RSof("NU_Peso")
			
		if mae=false AND fil =false AND in_co=True AND isnull(nu_peso) then
		
		verifica="ok"
		
			
				Set RSFIL = Server.CreateObject("ADODB.Recordset")
				SQLFIL = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"' AND CO_Materia_Principal='"&materia_fil&"'" 
				RSFIL.Open SQLFIL, CON2
		
			notaFIL=RSFIL("TP_Nota")
			
		
		if notaFIL ="TB_NOTA_A" then
		CAMINHOn = CAMINHO_na
		
		elseif notaFIL="TB_NOTA_B" then
			CAMINHOn = CAMINHO_nb
		
		elseif notaFIL ="TB_NOTA_C" then
				CAMINHOn = CAMINHO_nc
				
		elseif notaFIL ="TB_NOTA_D" then
				CAMINHOn = CAMINHO_nd
				
		elseif notaFIL ="TB_NOTA_E" then
				CAMINHOn = CAMINHO_ne
												
		else
				response.Write("ERRO")
		end if	
		
				Set CONn = Server.CreateObject("ADODB.Connection") 
				ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
				CONn.Open ABRIRn
		
		periodofil = periodo
		
				Set RSnFIL = Server.CreateObject("ADODB.Recordset")
				SQLnFIL = "SELECT * FROM "&notaFIL&" where CO_Matricula ="& codigo &" AND CO_Materia ='"& materia_fil &"' AND CO_Materia_Principal ='"& materia_fil &"' AND NU_Periodo="&periodofil
				RSnFIL.Open SQLnFIL, CONn
		
			Set RSPESO = Server.CreateObject("ADODB.Recordset")
				SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodofil
				RSPESO.Open SQLPESO, CON0
				
		
		if RSnFIL.EOF then
			n1= "&nbsp;"
			r1= "&nbsp;"
			f1= "&nbsp;"
			m1= "&nbsp;"
			peso1=0
			sm1=0
			contanotas=contanotas
			div_princ1=div_princ
			nota_princ1=nota_princ1
			ACU_Peso1=ACU_Peso1
			ACU_Rec1=ACU_Rec1
		else
			n1=RSnFIL("VA_Media2")
			r1=RSnFIL("VA_Rec")
			f1=RSnFIL("NU_Faltas")
			m1=RSnFIL("VA_Media3")
			peso1=RSPESO("NU_Peso")
			sm1=m1
			contanotas=contanotas+1
			div_princ1 = div_princ1+1
			nota_princ1=nota_princ1 + (n1*nu_peso2) 
			ACU_Peso1=ACU_Peso1+nu_peso2
			
			if isnull(r1) or r1="" then
				acu_r1=acu_r1
				ACU_Rec1=ACU_Rec1
			else
				acu_r1=acu_r1*1
				r1=r1*1
				acu_r1=acu_r1+r1
				ACU_Rec1=ACU_Rec1+1
			end if
		
		end if
		
		mp1 = nota_princ1
		mrec1=acu_r1
		
		if mp1>=mrec1 then
		media1=mp1
		else
		media1=(mp1+mrec1)/2
		end if
		
								decimo = media1 - Int(media1)
									If decimo >= 0.5 Then
									nota_arredondada = Int(media1) + 1
									mp1=nota_arredondada
									Else
									nota_arredondada = Int(media1)
									media1=nota_arredondada					
									End If
								media1 = formatNumber(media1,0)
							
								
		media1=media1*1
		ordem2=ordem2+1
		
		else
		verifica = "erro"
		END IF
	wend







elseif mae=TRUE AND fil=true AND in_co=false AND isnull(nu_peso) then

	Set RS1b = Server.CreateObject("ADODB.Recordset")
	SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&materia&"'"
	RS1b.Open SQL1b, CON0

	nota_princ1=0
	nota_princ2=0
	nota_princ3=0
	nota_princ4=0
	divisor=0
	divisor2=0
	
	contanotas=0
	div_princ1=0
	
	sm1=0
	sm2=0
	sm3=0
	sm4=0
	
	media1=0
	media2=0
	media3=0
	somamp=0
	nu_peso2=0
	mamp=0
	divisormp=0
	
	ACU_Rec1=0
	ACU_Rec2=0
	ACU_Rec3=0
	ACU_Rec4=0
	
	ACU_Peso1=0
	ACU_Peso2=0
	ACU_Peso3=0
	ACU_Peso4=0
	
	acu_r1=0
	acu_r2=0
	acu_r3=0
	acu_r4=0
	
	mr1=0
	mr2=0
	mr3=0
	mr4=0
	
	mp1=0
	mp2=0
	mp3=0
	mp4=0

	'check=check+1
	nota_princ1 = 0
	nota_princ2 = 0
	nota_princ3 = 0
	nota_princ4 = 0
	div_princ1 = 0
	div_princ2 = 0
	div_princ3 = 0
	div_princ4 = 0
	mat_fil_check="primeiro_"
	ACU_Peso1 = 0
	While not RS1b.EOF
		materia_fil =RS1b("CO_Materia")
		
		Set RS1t = Server.CreateObject("ADODB.Recordset")
		SQL1t = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia='"&materia_fil&"' order by NU_Ordem_Boletim "
		RS1t.Open SQL1t, CON0
			
	
		mae2=RS1t("IN_MAE")
		fil2=RS1t("IN_FIL")
		in_co2=RS1t("IN_CO")
		nu_peso2=RS1t("NU_Peso")
	'response.Write("nota_princ1 - "&nu_peso2&"-"&fil2&"-"&mae2&"-"&in_co2&"<br>")
		if mae2=false AND fil2 =true AND in_co2=false then		
		
			
			Set RS1c = Server.CreateObject("ADODB.Recordset")
			SQL1c = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia_fil&"'"
			RS1c.Open SQL1c, CON0
				
			no_materia_fil=RS1b("NO_Materia")
			
		
			Set RSFIL = Server.CreateObject("ADODB.Recordset")
			SQLFIL = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"' AND CO_Materia_Principal='"&materia_fil&"'" 
			RSFIL.Open SQLFIL, CON2
		
		
			co_mat_check= "nulo"
				while not RSFIL.EOF
					co_materia_teste = RSFIL("CO_Materia_Principal")
					if co_materia_teste = co_mat_check then
						RSFIL.MOVENEXT
					else
						contanotas=0
						divisor=0
						notaFIL=RSFIL("TP_Nota")
						prof_fil=RSFIL("CO_Professor")
						
						if notaFIL ="TB_NOTA_A" then
						CAMINHOn = CAMINHO_na
						
						elseif notaFIL="TB_NOTA_B" then
							CAMINHOn = CAMINHO_nb
						
						elseif notaFIL ="TB_NOTA_C" then
								CAMINHOn = CAMINHO_nc

						elseif notaFIL ="TB_NOTA_D" then
								CAMINHOn = CAMINHO_nd
								
						elseif notaFIL ="TB_NOTA_E" then
								CAMINHOn = CAMINHO_ne
				
						else
								response.Write("ERRO")
						end if	
		
						Set CONn = Server.CreateObject("ADODB.Connection") 
						ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
						CONn.Open ABRIRn
		
						periodofil = periodo
				
						Set RSnFIL = Server.CreateObject("ADODB.Recordset")
						SQLnFIL = "SELECT * FROM "&notaFIL&" where CO_Matricula ="& codigo &" AND CO_Materia ='"& materia_fil &"' AND CO_Materia_Principal ='"& materia &"' AND NU_Periodo="&periodofil
						RSnFIL.Open SQLnFIL, CONn
				
						Set RSPESO = Server.CreateObject("ADODB.Recordset")
						SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodofil
						RSPESO.Open SQLPESO, CON0
		
						if RSnFIL.EOF then
							n1= "&nbsp;"
							r1= "&nbsp;"
							f1= "&nbsp;"
							m1= "&nbsp;"
							peso1=0
							sm1=0
							contanotas=contanotas
							div_princ1=div_princ
							nota_princ1=nota_princ1
							ACU_Peso1=ACU_Peso1
						else
							n1=RSnFIL("VA_Media2")
							r1=RSnFIL("VA_Rec")
							f1=RSnFIL("NU_Faltas")
							m1=RSnFIL("VA_Media3")
							peso1=RSPESO("NU_Peso")
							sm1=m1
							contanotas=contanotas+1
							div_princ1 = div_princ1+1
							nota_princ1=nota_princ1 + (m1*nu_peso2) 
							ACU_Peso1=ACU_Peso1+peso1
						
						if isnull(r1) or r1="" then
							acu_r1=acu_r1
							ACU_Rec1=ACU_Rec1+1
						else
							acu_r1=acu_r1*1
							r1=r1*1
							acu_r1=acu_r1+r1
							ACU_Rec1=ACU_Rec1+1
						end if						
					end if
						
					m1=sm1*peso1						
					'check=check+1
					RSFIL.MOVENEXT
					co_mat_check = co_materia_teste
				end if
			wend				  
		end if
		RS1b.MOVENEXT
	wend
	
	if ACU_Peso1 = 0 then
		mp1 = nota_princ1
	else
		mp1 = nota_princ1 / ACU_Peso1
	end if
	
	'if ACU_Rec1=0 then
	'	mr1=0
	'else
	'	mr1 = acu_r1/ACU_Rec1
	'end if
	
	'if mr1>mp1 then
	'	media1=mr1
	'else
		media1=mp1
	'end if
	
	
	
	
	if media1="" or isnull(media1) then
	else
							decimo = media1 - Int(media1)
								If decimo >= 0.5 Then
								nota_arredondada = Int(media1) + 1
								media1=nota_arredondada
								Else
								nota_arredondada = Int(media1)
								media1=nota_arredondada					
								End If
							media1 = formatNumber(media1,0)
							
	media1=media1*1
	end if

elseif mae=TRUE AND fil=true AND in_co=false then
		Set RS1b = Server.CreateObject("ADODB.Recordset")
		SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&materia&"'"
		RS1b.Open SQL1b, CON0


		Set RSFIL0 = Server.CreateObject("ADODB.Recordset")
SQLFIL0 = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"' AND CO_Materia_Principal='"&materia&"'" 
		RSFIL0.Open SQLFIL0, CON2

		Set RS1t = Server.CreateObject("ADODB.Recordset")
		SQL1t = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia='"&materia&"' order by NU_Ordem_Boletim "
		RS1t.Open SQL1t, CON0
		
				
'RESPONSE.Write("SQL1t = SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia='"&materia&"' order by NU_Ordem_Boletim ")

		
nu_peso2=RS1t("NU_Peso")

	notaFIL=RSFIL0("TP_Nota")
if notaFIL ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na

elseif notaFIL="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb

elseif notaFIL ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc

elseif notaFIL ="TB_NOTA_D" then
		CAMINHOn = CAMINHO_nd
		
elseif notaFIL ="TB_NOTA_E" then
		CAMINHOn = CAMINHO_ne

else
		response.Write("ERRO")
end if	

		Set CONn = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONn.Open ABRIRn

periodofil = periodo
		
		Set RSnFIL = Server.CreateObject("ADODB.Recordset")
		SQLnFIL = "SELECT * FROM "&notaFIL&" where CO_Matricula ="& codigo &" AND CO_Materia ='"& materia &"' AND CO_Materia_Principal ='"& materia &"' AND NU_Periodo="&periodofil
		RSnFIL.Open SQLnFIL, CONn


		Set RSPESO = Server.CreateObject("ADODB.Recordset")
		SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodofil
		RSPESO.Open SQLPESO, CON0


if RSnFIL.EOF then
n1= ""
r1= ""
f1= ""
m1= ""
peso1=0
sm1=0
contanotas=contanotas
div_princ1=div_princ
nota_princ1=nota_princ1
ACU_Peso1=ACU_Peso1
ACU_Rec1=ACU_Rec1
else
n1=RSnFIL("VA_Media2")
r1=RSnFIL("VA_Rec")
f1=RSnFIL("NU_Faltas")
m1=RSnFIL("VA_Media3")
peso1=RSPESO("NU_Peso")
sm1=m1
contanotas=contanotas+1
div_princ1 = div_princ1+1
nota_princ1=nota_princ1 + (m1*nu_peso2) 
ACU_Peso1=ACU_Peso1+nu_peso2

if isnull(r1) or r1="" then
acu_r1=acu_r1
ACU_Rec1=ACU_Rec1+1
else
acu_r1=acu_r1*1
r1=r1*1
acu_r1=acu_r1+r1
ACU_Rec1=ACU_Rec1+1
end if

end if

m1=sm1*peso1
'check=check+1
mat_fil_check="primeiro_"
While not RS1b.EOF
	materia_fil =RS1b("CO_Materia")
	
		Set RS1t = Server.CreateObject("ADODB.Recordset")
		SQL1t = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia='"&materia_fil&"' order by NU_Ordem_Boletim "
		RS1t.Open SQL1t, CON0
		

	mae2=RS1t("IN_MAE")
	fil2=RS1t("IN_FIL")
	in_co2=RS1t("IN_CO")
	nu_peso2=RS1t("NU_Peso")
'response.Write("nota_princ1 - "&nu_peso2&"-"&fil2&"-"&mae2&"-"&in_co2&"<br>")
if mae2=false AND fil2 =true AND in_co2=false then		

	
		Set RS1c = Server.CreateObject("ADODB.Recordset")
		SQL1c = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia_fil&"'"
		RS1c.Open SQL1c, CON0
		
	no_materia_fil=RS1b("NO_Materia")
	


		Set RSFIL = Server.CreateObject("ADODB.Recordset")
		SQLFIL = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"' AND CO_Materia_Principal='"&materia_fil&"'" 
		RSFIL.Open SQLFIL, CON2


co_mat_check= "nulo"
while not RSFIL.EOF
co_materia_teste = RSFIL("CO_Materia_Principal")
if co_materia_teste = co_mat_check then
RSFIL.MOVENEXT
else
contanotas=0
divisor=0
	notaFIL=RSFIL("TP_Nota")
	prof_fil=RSFIL("CO_Professor")
if notaFIL ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na

elseif notaFIL="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb

elseif notaFIL ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc

elseif notaFIL ="TB_NOTA_D" then
		CAMINHOn = CAMINHO_nd
		
elseif notaFIL ="TB_NOTA_E" then
		CAMINHOn = CAMINHO_ne
			
else
		response.Write("ERRO")
end if	

		Set CONn = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONn.Open ABRIRn



periodofil = periodo
		
		Set RSnFIL = Server.CreateObject("ADODB.Recordset")
		SQLnFIL = "SELECT * FROM "&notaFIL&" where CO_Matricula ="& codigo &" AND CO_Materia ='"& materia_fil &"' AND CO_Materia_Principal ='"& materia &"' AND NU_Periodo="&periodofil
		RSnFIL.Open SQLnFIL, CONn


		Set RSPESO = Server.CreateObject("ADODB.Recordset")
		SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodofil
		RSPESO.Open SQLPESO, CON0


if RSnFIL.EOF then
n1= ""
r1= ""
f1= ""
m1= ""
peso1=0
sm1=0
contanotas=contanotas
div_princ1=div_princ
nota_princ1=nota_princ1
ACU_Peso1=ACU_Peso1
ACU_Rec1=ACU_Rec1
else
n1=RSnFIL("VA_Media2")
r1=RSnFIL("VA_Rec")
f1=RSnFIL("NU_Faltas")
m1=RSnFIL("VA_Media3")
peso1=RSPESO("NU_Peso")
sm1=m1
contanotas=contanotas+1
div_princ1 = div_princ1+1
nota_princ1=nota_princ1 + (m1*nu_peso2) 
ACU_Peso1=ACU_Peso1+nu_peso2

if isnull(r1) or r1="" then
acu_r1=acu_r1
ACU_Rec1=ACU_Rec1+1
else
acu_r1=acu_r1*1
r1=r1*1
acu_r1=acu_r1+r1
ACU_Rec1=ACU_Rec1+1
end if

end if



m1=sm1*peso1

'check=check+1
RSFIL.MOVENEXT
co_mat_check = co_materia_teste
end if
wend
		  
end if
RS1b.MOVENEXT
wend



if ACU_Peso1 = 0 then
mp1 = nota_princ1
else
mp1 = nota_princ1 / ACU_Peso1
end if

if ACU_Rec1=0 then
mr1=0
else
mr1 = acu_r1/ACU_Rec1
end if

if mr1>mp1 then
media1=mr1
else
media1=mp1
end if





						decimo = media1 - Int(media1)
							If decimo >= 0.5 Then
							nota_arredondada = Int(media1) + 1
							media1=nota_arredondada
							Else
							nota_arredondada = Int(media1)
							media1=nota_arredondada					
							End If
						media1 = formatNumber(media1,0)
						
				
media1=media1*1


END IF
                      
response.Write(media1)
END Function

























FUNCTION calculamediarec(codigo,unidade,curso,co_etapa,turma,materia,periodo,minimo)
' MCAL

contanotas=0
verifica = "ok"

ACU_Peso1 = 0
ACU_Peso2 = 0
ACU_Peso3 = 0
ACU_Peso4 = 0
nota_princ1=0
nota_princ2=0
nota_princ3=0
nota_princ4=0			  
nu_peso2 = 1

ACU_Rec1=0
ACU_Rec2=0
ACU_Rec3=0
ACU_Rec4=0

acu_r1=0
acu_r2=0
acu_r3=0
acu_r4=0

mrec1=0
mrec2=0
mrec3=0
mrec4=0

min=minimo


'response.Write("SQLo = SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' And CO_Materia ='"& materia&"'" )
		
		Set RSo = Server.CreateObject("ADODB.Recordset")
		SQLo = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' And CO_Materia ='"& materia&"'" 
		RSo.Open SQLo, CON0
		
	mae=RSo("IN_MAE")
	fil=RSo("IN_FIL")
	in_co=RSo("IN_CO")
	nu_peso=RSo("NU_Peso")
	ordem=RSo("NU_Ordem_Boletim")

'MCAL		
if mae=TRUE AND fil=false AND in_co=true AND isnull(nu_peso) then
	ordem2=ordem+1
While verifica = "ok"

'response.Write(">>"&ordem2)

		Set RSof = Server.CreateObject("ADODB.Recordset")
		SQLof = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' And NU_Ordem_Boletim="&ordem2 
		RSof.Open SQLof, CON0
	
	materia_fil=RSof("CO_Materia")
	mae=RSof("IN_MAE")
	fil=RSof("IN_FIL")
	in_co=RSof("IN_CO")
	nu_peso=RSof("NU_Peso")
	
if mae=false AND fil =false AND in_co=True AND isnull(nu_peso) then

verifica="ok"

	
		Set RSFIL = Server.CreateObject("ADODB.Recordset")
		SQLFIL = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"' AND CO_Materia_Principal='"&materia_fil&"'" 
		RSFIL.Open SQLFIL, CON2

	notaFIL=RSFIL("TP_Nota")
	

if notaFIL ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na

elseif notaFIL="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb

elseif notaFIL ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc

elseif notaFIL ="TB_NOTA_D" then
		CAMINHOn = CAMINHO_nd
		
elseif notaFIL ="TB_NOTA_E" then
		CAMINHOn = CAMINHO_ne

else
		response.Write("ERRO")
end if	

		Set CONn = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONn.Open ABRIRn

periodofil = periodo

		Set RSnFIL = Server.CreateObject("ADODB.Recordset")
		SQLnFIL = "SELECT * FROM "&notaFIL&" where CO_Matricula ="& codigo &" AND CO_Materia ='"& materia_fil &"' AND CO_Materia_Principal ='"& materia_fil &"' AND NU_Periodo="&periodofil
		RSnFIL.Open SQLnFIL, CONn

	Set RSPESO = Server.CreateObject("ADODB.Recordset")
		SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodofil
		RSPESO.Open SQLPESO, CON0
		

if RSnFIL.EOF then
n1= "&nbsp;"
r1= "&nbsp;"
f1= "&nbsp;"
m1= "&nbsp;"
peso1=0
sm1=0
contanotas=contanotas
div_princ1=div_princ
nota_princ1=nota_princ1
ACU_Peso1=ACU_Peso1
ACU_Rec1=ACU_Rec1
else
n1=RSnFIL("VA_Media2")
r1=RSnFIL("VA_Rec")
f1=RSnFIL("NU_Faltas")
m1=RSnFIL("VA_Media3")
peso1=RSPESO("NU_Peso")
sm1=m1
contanotas=contanotas+1
div_princ1 = div_princ1+1
nota_princ1=nota_princ1 + (n1*nu_peso2) 
ACU_Peso1=ACU_Peso1+nu_peso2

if isnull(r1) or r1="" then
acu_r1=acu_r1
ACU_Rec1=ACU_Rec1
else
acu_r1=acu_r1*1
r1=r1*1
acu_r1=acu_r1+r1
ACU_Rec1=ACU_Rec1+1
end if

end if

mp1 = nota_princ1
mrec1=acu_r1

if mp1>=mrec1 then
media1=mp1
else
media1=(mp1+mrec1)/2
end if

						decimo = media1 - Int(media1)
							If decimo >= 0.5 Then
							nota_arredondada = Int(media1) + 1
							mp1=nota_arredondada
							Else
							nota_arredondada = Int(media1)
							media1=nota_arredondada					
							End If
						media1 = formatNumber(media1,0)
					
						
media1=media1*1
ordem2=ordem2+1

else
verifica = "erro"
END IF
wend







elseif mae=TRUE AND fil=true AND in_co=false AND isnull(nu_peso) then

Set RS1b = Server.CreateObject("ADODB.Recordset")
		SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&materia&"'"
		RS1b.Open SQL1b, CON0

nota_princ1=0
nota_princ2=0
nota_princ3=0
nota_princ4=0
divisor=0
divisor2=0

contanotas=0
div_princ1=0

sm1=0
sm2=0
sm3=0
sm4=0

media1=0
media2=0
media3=0
somamp=0
nu_peso2=0
mamp=0
divisormp=0

ACU_Rec1=0
ACU_Rec2=0
ACU_Rec3=0
ACU_Rec4=0

ACU_Peso1=0
ACU_Peso2=0
ACU_Peso3=0
ACU_Peso4=0

acu_r1=0
acu_r2=0
acu_r3=0
acu_r4=0

mr1=0
mr2=0
mr3=0
mr4=0

mp1=0
mp2=0
mp3=0
mp4=0

'check=check+1
nota_princ1 = 0
nota_princ2 = 0
nota_princ3 = 0
nota_princ4 = 0
div_princ1 = 0
div_princ2 = 0
div_princ3 = 0
div_princ4 = 0
mat_fil_check="primeiro_"
ACU_Peso1 = 0
While not RS1b.EOF
	materia_fil =RS1b("CO_Materia")
	
		Set RS1t = Server.CreateObject("ADODB.Recordset")
		SQL1t = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia='"&materia_fil&"' order by NU_Ordem_Boletim "
		RS1t.Open SQL1t, CON0
		

	mae2=RS1t("IN_MAE")
	fil2=RS1t("IN_FIL")
	in_co2=RS1t("IN_CO")
	nu_peso2=RS1t("NU_Peso")
'response.Write("nota_princ1 - "&nu_peso2&"-"&fil2&"-"&mae2&"-"&in_co2&"<br>")
if mae2=false AND fil2 =true AND in_co2=false then		

	
		Set RS1c = Server.CreateObject("ADODB.Recordset")
		SQL1c = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia_fil&"'"
		RS1c.Open SQL1c, CON0
		
	no_materia_fil=RS1b("NO_Materia")
	

		Set RSFIL = Server.CreateObject("ADODB.Recordset")
		SQLFIL = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"' AND CO_Materia_Principal='"&materia_fil&"'" 
		RSFIL.Open SQLFIL, CON2


co_mat_check= "nulo"
while not RSFIL.EOF
co_materia_teste = RSFIL("CO_Materia_Principal")
if co_materia_teste = co_mat_check then
RSFIL.MOVENEXT
else
contanotas=0
divisor=0
	notaFIL=RSFIL("TP_Nota")
	prof_fil=RSFIL("CO_Professor")
if notaFIL ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na

elseif notaFIL="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb

elseif notaFIL ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc

elseif notaFIL ="TB_NOTA_D" then
		CAMINHOn = CAMINHO_nd
		
elseif notaFIL ="TB_NOTA_E" then
		CAMINHOn = CAMINHO_ne

else
		response.Write("ERRO")
end if	

		Set CONn = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONn.Open ABRIRn

periodofil = periodo
		
		Set RSnFIL = Server.CreateObject("ADODB.Recordset")
		SQLnFIL = "SELECT * FROM "&notaFIL&" where CO_Matricula ="& codigo &" AND CO_Materia ='"& materia_fil &"' AND CO_Materia_Principal ='"& materia &"' AND NU_Periodo="&periodofil
		RSnFIL.Open SQLnFIL, CONn


		Set RSPESO = Server.CreateObject("ADODB.Recordset")
		SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodofil
		RSPESO.Open SQLPESO, CON0



if RSnFIL.EOF then
n1= "&nbsp;"
r1= "&nbsp;"
f1= "&nbsp;"
m1= "&nbsp;"
peso1=0
sm1=0
contanotas=contanotas
div_princ1=div_princ
nota_princ1=nota_princ1
ACU_Peso1=ACU_Peso1
else
n1=RSnFIL("VA_Media2")
r1=RSnFIL("VA_Rec")
f1=RSnFIL("NU_Faltas")
m1=RSnFIL("VA_Media3")
peso1=RSPESO("NU_Peso")
sm1=m1
contanotas=contanotas+1
div_princ1 = div_princ1+1
nota_princ1=nota_princ1 + (n1*nu_peso2) 
ACU_Peso1=ACU_Peso1+nu_peso2

if isnull(r1) or r1="" then
acu_r1=acu_r1
ACU_Rec1=ACU_Rec1+1
else
acu_r1=acu_r1*1
r1=r1*1
acu_r1=acu_r1+r1
ACU_Rec1=ACU_Rec1+1
end if

end if



m1=sm1*peso1

'check=check+1
RSFIL.MOVENEXT
co_mat_check = co_materia_teste
end if
wend
		  
end if
RS1b.MOVENEXT
wend

if ACU_Peso1 = 0 then
mp1 = nota_princ1
else
mp1 = nota_princ1 / ACU_Peso1
end if

if ACU_Rec1=0 then
mr1=0
else
mr1 = acu_r1/ACU_Rec1
end if

if mr1>mp1 then
media1=mr1
else
media1=mp1
end if





						decimo = media1 - Int(media1)
							If decimo >= 0.5 Then
							nota_arredondada = Int(media1) + 1
							media1=nota_arredondada
							Else
							nota_arredondada = Int(media1)
							media1=nota_arredondada					
							End If
						media1 = formatNumber(media1,0)
						
media1=media1*1










elseif mae=TRUE AND fil=true AND in_co=false then
		Set RS1b = Server.CreateObject("ADODB.Recordset")
		SQL1b = "SELECT * FROM TB_Materia WHERE CO_Materia_Principal='"&materia&"'"
		RS1b.Open SQL1b, CON0


		Set RSFIL0 = Server.CreateObject("ADODB.Recordset")
SQLFIL0 = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"' AND CO_Materia_Principal='"&materia&"'" 
		RSFIL0.Open SQLFIL0, CON2

		Set RS1t = Server.CreateObject("ADODB.Recordset")
		SQL1t = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia='"&materia&"' order by NU_Ordem_Boletim "
		RS1t.Open SQL1t, CON0
		
				
'RESPONSE.Write("SQL1t = SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia='"&materia&"' order by NU_Ordem_Boletim ")

		
nu_peso2=RS1t("NU_Peso")

	notaFIL=RSFIL0("TP_Nota")
if notaFIL ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na

elseif notaFIL="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb

elseif notaFIL ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc
		
elseif notaFIL ="TB_NOTA_D" then
		CAMINHOn = CAMINHO_nd
		
elseif notaFIL ="TB_NOTA_E" then
		CAMINHOn = CAMINHO_ne
	
else
		response.Write("ERRO")
end if	

		Set CONn = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONn.Open ABRIRn

periodofil = periodo
		
		Set RSnFIL = Server.CreateObject("ADODB.Recordset")
		SQLnFIL = "SELECT * FROM "&notaFIL&" where CO_Matricula ="& codigo &" AND CO_Materia ='"& materia &"' AND CO_Materia_Principal ='"& materia &"' AND NU_Periodo="&periodofil
		RSnFIL.Open SQLnFIL, CONn


		Set RSPESO = Server.CreateObject("ADODB.Recordset")
		SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodofil
		RSPESO.Open SQLPESO, CON0


if RSnFIL.EOF then
n1= ""
r1= ""
f1= ""
m1= ""
peso1=0
sm1=0
contanotas=contanotas
div_princ1=div_princ
nota_princ1=nota_princ1
ACU_Peso1=ACU_Peso1
ACU_Rec1=ACU_Rec1
else
n1=RSnFIL("VA_Media2")
r1=RSnFIL("VA_Rec")
f1=RSnFIL("NU_Faltas")
m1=RSnFIL("VA_Media3")
peso1=RSPESO("NU_Peso")
sm1=m1
contanotas=contanotas+1
div_princ1 = div_princ1+1
nota_princ1=nota_princ1 + (m1*nu_peso2) 
ACU_Peso1=ACU_Peso1+nu_peso2

if isnull(r1) or r1="" then
acu_r1=acu_r1
ACU_Rec1=ACU_Rec1+1
else
acu_r1=acu_r1*1
r1=r1*1
acu_r1=acu_r1+r1
ACU_Rec1=ACU_Rec1+1
end if

end if

m1=sm1*peso1
'check=check+1
mat_fil_check="primeiro_"
While not RS1b.EOF
	materia_fil =RS1b("CO_Materia")
	
		Set RS1t = Server.CreateObject("ADODB.Recordset")
		SQL1t = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia='"&materia_fil&"' order by NU_Ordem_Boletim "
		RS1t.Open SQL1t, CON0
		

	mae2=RS1t("IN_MAE")
	fil2=RS1t("IN_FIL")
	in_co2=RS1t("IN_CO")
	nu_peso2=RS1t("NU_Peso")
'response.Write("nota_princ1 - "&nu_peso2&"-"&fil2&"-"&mae2&"-"&in_co2&"<br>")
if mae2=false AND fil2 =true AND in_co2=false then		

	
		Set RS1c = Server.CreateObject("ADODB.Recordset")
		SQL1c = "SELECT * FROM TB_Materia WHERE CO_Materia='"&materia_fil&"'"
		RS1c.Open SQL1c, CON0
		
	no_materia_fil=RS1b("NO_Materia")
	


		Set RSFIL = Server.CreateObject("ADODB.Recordset")
		SQLFIL = "SELECT * FROM TB_Da_Aula where NU_Unidade ="& unidade &" AND CO_Curso ='"& curso &"' AND CO_Etapa ='"& co_etapa &"' AND CO_Turma ='"& turma &"' AND CO_Materia_Principal='"&materia_fil&"'" 
		RSFIL.Open SQLFIL, CON2


co_mat_check= "nulo"
while not RSFIL.EOF
co_materia_teste = RSFIL("CO_Materia_Principal")
if co_materia_teste = co_mat_check then
RSFIL.MOVENEXT
else
contanotas=0
divisor=0
	notaFIL=RSFIL("TP_Nota")
	prof_fil=RSFIL("CO_Professor")
if notaFIL ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na

elseif notaFIL="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb

elseif notaFIL ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc
		
elseif notaFIL ="TB_NOTA_D" then
		CAMINHOn = CAMINHO_nd
		
elseif notaFIL ="TB_NOTA_E" then
		CAMINHOn = CAMINHO_ne

else
		response.Write("ERRO")
end if	

		Set CONn = Server.CreateObject("ADODB.Connection") 
		ABRIRn = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONn.Open ABRIRn



periodofil = periodo
		
		Set RSnFIL = Server.CreateObject("ADODB.Recordset")
		SQLnFIL = "SELECT * FROM "&notaFIL&" where CO_Matricula ="& codigo &" AND CO_Materia ='"& materia_fil &"' AND CO_Materia_Principal ='"& materia &"' AND NU_Periodo="&periodofil
		RSnFIL.Open SQLnFIL, CONn


		Set RSPESO = Server.CreateObject("ADODB.Recordset")
		SQLPESO = "SELECT * FROM TB_Periodo where NU_Periodo ="&periodofil
		RSPESO.Open SQLPESO, CON0


if RSnFIL.EOF then
n1= ""
r1= ""
f1= ""
m1= ""
peso1=0
sm1=0
contanotas=contanotas
div_princ1=div_princ
nota_princ1=nota_princ1
ACU_Peso1=ACU_Peso1
ACU_Rec1=ACU_Rec1
else
n1=RSnFIL("VA_Media2")
r1=RSnFIL("VA_Rec")
f1=RSnFIL("NU_Faltas")
m1=RSnFIL("VA_Media3")
peso1=RSPESO("NU_Peso")
sm1=m1
contanotas=contanotas+1
div_princ1 = div_princ1+1
nota_princ1=nota_princ1 + (m1*nu_peso2) 
ACU_Peso1=ACU_Peso1+nu_peso2

if isnull(r1) or r1="" then
acu_r1=acu_r1
ACU_Rec1=ACU_Rec1+1
else
acu_r1=acu_r1*1
r1=r1*1
acu_r1=acu_r1+r1
ACU_Rec1=ACU_Rec1+1
end if

end if



m1=sm1*peso1

'check=check+1
RSFIL.MOVENEXT
co_mat_check = co_materia_teste
end if
wend
		  
end if
RS1b.MOVENEXT
wend



if ACU_Peso1 = 0 then
mp1 = nota_princ1
else
mp1 = nota_princ1 / ACU_Peso1
end if

if ACU_Rec1=0 then
mr1=0
else
mr1 = acu_r1/ACU_Rec1
end if

if mr1>mp1 then
media1=mr1
else
media1=mp1
end if





						decimo = media1 - Int(media1)
							If decimo >= 0.5 Then
							nota_arredondada = Int(media1) + 1
							media1=nota_arredondada
							Else
							nota_arredondada = Int(media1)
							media1=nota_arredondada					
							End If
						media1 = formatNumber(media1,0)
						
				
media1=media1*1


END IF
					  'response.Write(media1 &">="& min )

					  if media1 >= min then
					  response.Write("&nbsp;")
					  elseif periodofil = 4 and media1=0 then
					  response.Write("&nbsp;")
					  else					  
					  response.Write(media1)
					  end if

END Function
%>

