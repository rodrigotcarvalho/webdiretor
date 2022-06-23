<!--#include file="caminhos.asp"-->
<!--#include file="utils.asp"-->
<%
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0


Function ehFeriado(dia)

	Set RSFeriado = Server.CreateObject("ADODB.Recordset")
	SQLFeriado = "Select * from TB_Feriados WHERE  DA_Inicio>= #"& dia &"# AND DA_Termino<=#"&dia&"#"
	Set RSFeriado = CON0.Execute(SQLFeriado)
	
	if RSFeriado.EOF then		
		ehFeriado = "N"
	else
		ehFeriado = "S"		
	end if

End function

Function diasPeriodo(periodoConsulta)

	Set RSPeriodo = Server.CreateObject("ADODB.Recordset")
	SQLPeriodo = "Select * from TB_Periodo WHERE NU_Periodo= "&periodoConsulta
	Set RSPeriodo = CON0.Execute(SQLPeriodo)
	
	dataInicio = RSPeriodo("DA_Inicio_Periodo")
	dataFim = RSPeriodo("DA_Fim_Periodo")
	
	dataInicioVetor= split(dataInicio,"/")
	diaInicio=dataInicioVetor(0)
	mesInicio=dataInicioVetor(1)
	anoInicio=dataInicioVetor(2)		
	dataFimVetor= split(dataFim,"/")
	diaFim=dataFimVetor(0)
	mesFim=dataFimVetor(1)
	anoFim=dataFimVetor(2)	
	totalDias=0
	mes=mesInicio			
	if anoInicio=anoFim then
		while mes<=mesFim
			mes=mes*1
			mesFim=mesFim*1
			if mes<>mesInicio then
				diaMes=1
			else
				diaMes=diaInicio
			end if
			for dia = diaMes to qtdDiasMes(mes,anoInicio)	
				dia=dia*1
				diaFim=diaFim*1
				if mes<mesFim or (mes=mesFim and dia<=diaFim) then
					if totalDias=0 then
						vetorDiasPeriodo=dia&"/"&mes&"/"&anoInicio					
					else
						vetorDiasPeriodo=vetorDiasPeriodo&"#!#"&dia&"/"&mes&"/"&anoInicio						
					end if	
				end if		
				totalDias = totalDias+1						
			next
			if mes<12 then
				mes = mes+1'
			end if
		wend
	end if
	diasPeriodo=vetorDiasPeriodo
End function

Function diasPeriodoFormatado(periodoConsulta,separador,formato)

	vetorDiasPeriodo=diasPeriodo(periodoConsulta)
	vetorDiasPeriodoFormatado=split(vetorDiasPeriodo,"#!#")
	for dpf=0 to ubound(vetorDiasPeriodoFormatado)
		dataFormatadaVetor= split(vetorDiasPeriodoFormatado(dpf),"/")
		diaFormatado=dataFormatadaVetor(0)
		mesFormatado=dataFormatadaVetor(1)
		anoFormatado=dataFormatadaVetor(2)	
		
		if formato="DD/MM" then
			if dpf=0 then
				vetorDatasPeriodoFormatada = formataData(diaFormatado)&"/"&formataData(mesFormatado)
			else	
				vetorDatasPeriodoFormatada = vetorDatasPeriodoFormatada&separador&formataData(diaFormatado)&"/"&formataData(mesFormatado)			
			end if
		elseif formato="DD/MM/YYYY" then	
			if dpf=0 then
				vetorDatasPeriodoFormatada = formataData(diaFormatado)&"/"&formataData(mesFormatado)&"/"&anoFormatado
			else	
				vetorDatasPeriodoFormatada = vetorDatasPeriodoFormatada&separador&formataData(diaFormatado)&"/"&formataData(mesFormatado)&"/"&anoFormatado			
			end if		
		end if	
	next
	
	diasPeriodoFormatado=vetorDatasPeriodoFormatada

End function
	
%>
