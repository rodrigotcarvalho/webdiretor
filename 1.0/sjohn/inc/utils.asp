<%
Function ehBissexto(anoConsulta)

	if anoConsulta mod 4 = 0 then
		ehBissexto = "S"
	else
		ehBissexto = "N"	
	end if

End Function

Function qtdDiasMes(mesConsulta,anoConsulta)

Select Case mesConsulta
case 1	
	qtdDiasMes=31
case 3
	qtdDiasMes=31
case 5
	qtdDiasMes=31
case 7
	qtdDiasMes=31
case 8
	qtdDiasMes=31
case 10
	qtdDiasMes=31
case 12		
	qtdDiasMes=31
case 4	
	qtdDiasMes=30
case 6
	qtdDiasMes=30
case 9
	qtdDiasMes=30
case 2
	qtdDiasMes=28
end select

if mesConsulta = 2 then 
	if ehBissexto(anoConsulta)="S" then
		qtdDiasMes=29
	end if
end if



End Function

Function formataData(diaOuMes)
	diaOuMes=diaOuMes*1
	if diaOuMes<10 then
		formataData="0"&diaOuMes
	else
		formataData=diaOuMes	
	end if 
End Function

Function formata(entrada,formato)
	if formato="DD/MM/YYYY" then	
		vetorData=split(entrada,"/")
		diaFormatado=vetorData(0)
		mesFormatado=vetorData(1)
		anoFormatado=vetorData(2)	
		formata=formataData(diaFormatado)&"/"&formataData(mesFormatado)&"/"&anoFormatado		
	else
		formata=entrada
	end if	

 
End Function
%>