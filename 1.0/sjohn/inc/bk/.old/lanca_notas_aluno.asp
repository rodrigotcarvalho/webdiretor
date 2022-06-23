<!--#include file="../../global/tabelas_escolas.asp"-->
<!--#include file="funcoes6.asp"-->
<script language="JavaScript" type="text/JavaScript">
function mudar_cor_focus(celula){
   celula.style.backgroundColor="#D8FF9D"

}
function mudar_cor_blur_par(celula){
   celula.style.backgroundColor="#FFFFFF"
} 
function mudar_cor_blur_impar(celula){
   celula.style.backgroundColor="#FFFFE1"
} 
function mudar_cor_blur_erro(celula){
   celula.style.backgroundColor="#CC0000"
}  
function checksubmit()
{
<%
co_usr = session("co_user")
if notaFIL<>"TB_NOTA_V" then
%>
// if (document.nota.pt.value == "")
//  {    alert("Por favor digite um peso para os Testes!")
//    document.nota.pt.focus()
//    return false
//  }
//  if (isNaN(document.nota.pt.value))
//  {    alert("O peso dos Testes deve ser um número!")
//    document.nota.pt.focus()
//    return false
//  }  

<%
end if
if notaFIL="TB_NOTA_F" then%>
//    if (document.nota.pp1.value == "")
//  {    alert("Por favor digite um peso para a Prova1!")
//    document.nota.pp1.focus()
//    return false
//  }
//  if (isNaN(document.nota.pp1.value))
//  {    alert("O peso da Prova 1 deve ser um número!")
//    document.nota.pp1.focus()
//    return false
//  }
//    if (document.nota.pp2.value == "")
//  {    alert("Por favor digite um peso para a Prova 2!")
//    document.nota.pp2.focus()
//    return false
//  }  
//  if (isNaN(document.nota.pp2.value))
//  {    alert("O peso das Prova 2 deve ser um número!")
//    document.nota.pp2.focus()
//    return false
//  }  
<%

elseif notaFIL<>"TB_NOTA_V" then%>
//    if (document.nota.pp.value == "")
//  {    alert("Por favor digite um peso para as Provas!")
//    document.nota.pp.focus()
//    return false
//  }
//  if (isNaN(document.nota.pp.value))
//  {    alert("O peso das Provas deve ser um número!")
//    document.nota.pp.focus()
//    return false
//  }
<%end if%>  
  return true
}
function MM_callJS(jsStr) { //v2.0
  return eval(jsStr)
}
    function keyPressed(TB, e, max_right, max_bottom)  
    { 
        if (e.keyCode == 40 || e.keyCode == 13) { // arrow down 
            if (TB.split("c")[0] < max_bottom) 
            document.getElementById(eval(TB.split("c")[0] + '+1') + 'c' + TB.split("c")[1]).focus(); 
            if (TB.split("c")[0] == max_bottom) 
            document.getElementById(1 + 'c' + TB.split("c")[1]).focus();


        } 
  
        if (e.keyCode == 38) { // arrow up 
            if(TB.split("c")[0] > 1) 
            document.getElementById(eval(TB.split("c")[0] + '-1') + 'c' + TB.split("c")[1]).focus(); 
            if (TB.split("c")[0] == 1) 
            document.getElementById(max_bottom + 'c' + TB.split("c")[1]).focus(); 
		
        } 
  
        if (e.keyCode == 37) { // arrow left 
            if(TB.split("c")[1] > 1) 
            document.getElementById(TB.split("c")[0] + 'c' + eval(TB.split("c")[1] + '-1')).focus();             
            if (TB.split("c")[1] == 1) 
            document.getElementById(TB.split("c")[0] + 'c' + max_right).focus(); 

		}   
  
        if (e.keyCode == 39) { // arrow right 
            if(TB.split("c")[1] < max_right) 
            document.getElementById(TB.split("c")[0] + 'c' + eval(TB.split("c")[1] + '+1')).focus();  
            if (TB.split("c")[1] == max_right) 
            document.getElementById(TB.split("c")[0] + 'c' + 1).focus(); 

		}                  
    } 
function MM_goToURL() { //v3.0
  var i, args=MM_goToURL.arguments; document.MM_returnValue = false;
  for (i=0; i<(args.length-1); i+=2) eval(args[i]+".location='"+args[i+1]+"'");
}	
//-->
</script>
<%
chave = session("chave")
session("chave")=chave
split_chave=split(chave,"-")
sistema_origem=split_chave(0)
funcao_origem=split_chave(3)

if sistema_origem="WN" then
	endereco_origem="../wn/lancar/notas/mna/"
elseif sistema_origem="WA" then	
'	if funcao_origem="EPN" then
'		endereco_origem="../wa/professor/relatorio/epn/"
'	else
		endereco_origem="../wa/professor/cna/mna/"
'	end if
end if	



		Set CON_N = Server.CreateObject("ADODB.Connection")
		ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_N.Open ABRIR3
		
		'Só usado no caso da tabela V que utiliza também a F
		'--------------------------------------------------------------------
		Set CON_NF = Server.CreateObject("ADODB.Connection")
		ABRIR3F = "DBQ="& CAMINHO_nf & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_NF.Open ABRIR3F		
		'--------------------------------------------------------------------
			
		Set CON_wr = Server.CreateObject("ADODB.Connection") 
		ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_wr.Open ABRIR_wr
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0
		
		Set CON_AL = Server.CreateObject("ADODB.Connection") 
		ABRIR_AL = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON_AL.Open ABRIR_AL
		
		Set CONg = Server.CreateObject("ADODB.Connection") 
		ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONg.Open ABRIRg		

linha_tabela=0
ntvmla0= 59
ntvmlb0= 59
ntvmlc0= 69
ntvmla=ntvmla0
ntvmlb=ntvmlb0
ntvmlc=ntvmlc0

' 		Set RS0 = Server.CreateObject("ADODB.Recordset")
'		SQL_0 = "Select * from TB_Materia WHERE CO_Materia = '"& co_materia &"'"
'		Set RS0 = CON_0.Execute(SQL_0)
'
'mat_princ=RS0("CO_Materia_Principal")
'
'if mat_princ="" or isnull(mat_princ) then
'	mat_princ=co_materia
'end if
'
'
'		Set RS = Server.CreateObject("ADODB.Recordset")
'		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& unidade &" AND CO_Curso = '"& curso &"' AND CO_Etapa = '"& etapa &"' AND CO_Turma = '"& turma &"' AND CO_Materia_Principal = '"& co_materia &"'"
'		Set RS = CONg.Execute(CONEXAO)
'
'ST_Per_1 = RS("ST_Per_1")
'ST_Per_2 = RS("ST_Per_2")
'ST_Per_3 = RS("ST_Per_3")
'ST_Per_4 = RS("ST_Per_4")
'ST_Per_5 = RS("ST_Per_5")
'ST_Per_6 = RS("ST_Per_6")

'response.Write(">"&opcao&"<")

'0 tb&"#$#"&
'1 ln_pesos_cols&"#$#"&
'2 ln_pesos_vars&"#$#"&
'3 nm_pesos_vars&"#$#"&
'4 ln_nom_cols&"#$#"&
'5 nm_vars&"#$#"&
'6 nm_bd&"#$#"&
'7 vars_calc&"#$#"&
'8 action&"#$#"&
'9 notas_a_lancar&"#$#"&
'10 gera_pdf&"#$#"&
'11 ln_bol_av_cols&"#$#"&
'12 ln_bol_av_span&"#$#"&
'13 nm_bol_av_vars&"#$#"&
'14 ln_bol_av_vars&"#$#"&
'15 vars_bol_av_calc&"#$#"&
'16 legenda_bol_av&"#$#"&
'17 exibe_apr_pr_bol_av
nu_matricula = cod
session("matricula")=nu_matricula


    dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,outro)
	dados_separados=split(dados_tabela,"#$#")
	tb=dados_separados(0)
	ln_pesos_cols=dados_separados(1)
	ln_pesos_vars=dados_separados(2)
	nm_pesos_vars=dados_separados(3)
	ln_nom_cols=dados_separados(18)
	nm_vars=dados_separados(19)
	nm_bd=dados_separados(20)
	action=dados_separados(21)
	notas_a_lancar=dados_separados(22)

	linha_pesos=split(ln_pesos_cols,"#!#")
	linha_pesos_variaveis=split(ln_pesos_vars,"#!#")
	nome_pesos_variaveis=split(nm_pesos_vars,"#!#")
	linha_nome_colunas=split(ln_nom_cols,"#!#")
	nome_variaveis=split(nm_vars,"#!#")
	variaveis_bd=split(nm_bd,"#!#")
	
	nova_linha_pesos = "&nbsp;#!#"&linha_pesos(1)
	qtd_colunas=0	
	for lp=3 to ubound(linha_pesos)
		if linha_pesos(lp)<>"&nbsp;" then
			qtd_colunas=qtd_colunas+1
		end if
		nova_linha_pesos = nova_linha_pesos&"#!#"&linha_pesos(lp)	
	next
	nova_linha_pesos = nova_linha_pesos
	
	nova_linha_pesos_variaveis = "&nbsp;"
	for lpv=2 to ubound(linha_pesos_variaveis)
		nova_linha_pesos_variaveis = nova_linha_pesos_variaveis&"#!#"&linha_pesos_variaveis(lpv)	
	next	
	nova_linha_pesos_variaveis = nova_linha_pesos_variaveis 	
		
	nova_nome_pesos_variaveis = "&nbsp;"
	for npv=2 to ubound(linha_pesos_variaveis)
		nova_nome_pesos_variaveis = nova_nome_pesos_variaveis&"#!#"&nome_pesos_variaveis(npv)	
	next
	nova_nome_pesos_variaveis = nova_nome_pesos_variaveis
	
	nova_linha_nome_colunas = "Disciplinas"
	for lnc=2 to ubound(linha_nome_colunas)
		nova_linha_nome_colunas = nova_linha_nome_colunas&"#!#"&linha_nome_colunas(lnc)	
	next				
	nova_linha_nome_colunas = nova_linha_nome_colunas
	
	linha_pesos=split(nova_linha_pesos,"#!#")
	linha_pesos_variaveis=split(nova_linha_pesos_variaveis,"#!#")
	nome_pesos_variaveis=split(nova_nome_pesos_variaveis,"#!#")

	linha_nome_colunas=split(nova_linha_nome_colunas,"#!#")
	
	for nv=0 to ubound(nome_variaveis)
		if nv=0 then
			nova_nome_variaveis = nome_variaveis(nv)			
		else
			nova_nome_variaveis = nova_nome_variaveis&"#!#"&nome_variaveis(nv)	
		end if	
	next		
	nova_nome_variaveis = nova_nome_variaveis
	
	for vb=0 to ubound(variaveis_bd)
		if vb=0 then
			nova_variaveis_bd = variaveis_bd(vb)			
		else
			nova_variaveis_bd = nova_variaveis_bd&"#!#"&variaveis_bd(vb)	
		end if	
	next		
	nova_variaveis_bd = nova_variaveis_bd

	nome_variaveis=split(nova_nome_variaveis,"#!#")
	variaveis_bd=split(nova_variaveis_bd,"#!#")	

if subopcao="cln" then
	comunica="s" 
	opt="?opt=cln&obr="&obr&"&nota="&tb
	tipo="hidden"

elseif subopcao="imp" then
	comunica="s" 
	opt=""
	tipo="hidden"
elseif subopcao="blq" then
	comunica="s" 
	opt=""
	tipo="hidden"
else
	comunica="n" 
	opt=""
	tipo="text"
end if

if subopcao="imp" Then
	classe_peso = "tabelaTit"
	classe_subtit = "tabelaTit"

elseif errou="pt" or errou="pp" Then
	classe_peso = "tb_fundo_linha_erro"
else
	classe_peso = "tb_fundo_linha_peso"
	classe_subtit = "tb_subtit"
end if


qtd_colunas=UBOUND(linha_nome_colunas)+1+qtd_colunas
width_nom_disciplina="200"
width_else=(1000-200)/(qtd_colunas-2)

	Set RS5 = Server.CreateObject("ADODB.Recordset")
	SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' order by NU_Ordem_Boletim"
	RS5.Open SQL5, CON0
	co_materia_check=1
	IF RS5.EOF Then
		vetor_materia_exibe="nulo"
	else
		while not RS5.EOF
			co_mat_fil= RS5("CO_Materia")				
			if co_materia_check=1 then
				vetor_materia=co_mat_fil
			else
				vetor_materia=vetor_materia&"#!#"&co_mat_fil
			end if
			co_materia_check=co_materia_check+1			
					
		RS5.MOVENEXT
		wend	
'	 	response.Write(vetor_materia&"<BR>")
'		if session("ano_letivo") < ano_letivo_prog_aula then
			vetor_materia_exibe=programa_aula(vetor_materia, unidade, curso, etapa, "nulo")
'		else
'			vetor_materia_exibe=programa_aula_boletim_ficha(vetor_materia, unidade, curso, etapa, "nulo")			
'		end if
'		response.Write(vetor_materia_exibe)
	end if	
	
	co_materia_exibe=Split(vetor_materia_exibe,"#!#")	
	qtd_disciplinas=0
	For c=0 to ubound(co_materia_exibe)
			
		co_materia_teste = co_materia_exibe(c)
		
		if co_materia_teste<> "MED" then
		
			Set RSprog = Server.CreateObject("ADODB.Recordset")
			SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia = '"& co_materia_teste &"'"
			RSprog.Open SQLprog, CON0
	
			mae=RSprog("IN_MAE")
			fil=RSprog("IN_FIL")
				
			if mae=TRUE and fil= TRUE then	
				qtd_disciplinas = qtd_disciplinas
			else
				qtd_disciplinas = qtd_disciplinas+1			
			end if
		end if
	next			
%>
<form action="../../../../inc/grava_notas_aluno.asp" name="nota" method="post" onSubmit="return checksubmit()">
<table width="1000" border="0" cellspacing="0" cellpadding="0">
<%'if ubound(linha_pesos)>-1 then %>
<!--  <tr> -->
 <% 
 'for i= 0 to ubound(linha_pesos)
' 		if i=0 then
'			width=width_num_cham
'			align="center"
''		elseif i=1 then	
''			width=width_nom_aluno
''			align="left"			
'		else
'			width=width_else
'			align="center"
'		end if				
'					
'		if linha_pesos(i)="PESO" then			
'		
'			For co=0 to ubound(co_materia_exibe)
'			
'				co_materia = co_materia_exibe(co)
'				
'				if co_materia<> "MED" then
'		
'					Set RS1a = Server.CreateObject("ADODB.Recordset")
'					SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&co_materia&"'"
'					RS1a.Open SQL1a, CON0
'						
'					if RS1a.eof then
'						response.Write("Erro no peso - Busca pela disciplina")
'						response.End()
'					else
'						co_materia_principal=RS1a("CO_Materia_Principal")	
'						if isnull(co_materia_principal)	or co_materia_principal = "" then
'							co_materia_principal = co_materia
'						end if			
'					end if	
'					Set RSpeso = Server.CreateObject("ADODB.Recordset")
'					SQL_peso = "Select "&linha_pesos_variaveis(i)&" from "& tb &" WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& co_materia_principal &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
'					Set RSpeso = CON_N.Execute(SQL_peso)
'					
'					if valor_peso = "" or isnull(valor_peso) then
'						if RSpeso.EOF then
'							if tb= "TB_NOTA_F" then
'								if periodo<=4 then
'									if linha_pesos_variaveis(i) = "PE_Teste" then
'										valor_peso = 1
'									elseif linha_pesos_variaveis(i) = "PE_Prova1" then
'										valor_peso = 4
'									else
'										if periodo<4 then
'											valor_peso = 5	
'										else
'											valor_peso =""										
'										end if						 
'									end if
'								 else
'								   valor_peso =""	
'								 end if	
'							end if				
'						else	
'
'							valor_peso=RSpeso(""&linha_pesos_variaveis(i)&"")						
'							if (valor_peso = "" or isnull(valor_peso)) and tb= "TB_NOTA_F" then
'								if periodo<=4 then
'									if linha_pesos_variaveis(i) = "PE_Teste" then
'										valor_peso = 1
'									elseif linha_pesos_variaveis(i) = "PE_Prova1" then
'										valor_peso = 4
'									else
'										if periodo<4 then
'											valor_peso = 5	
'										else
'											valor_peso =""										
'										end if						 
'									end if
'								 else
'								   valor_peso =""	
'								 end if	
'							end if
'						end if	
'					end if
'				end if	
'			Next
'			IF comunica="s" THEN		
'				linha_pesos(i)=valor_peso&"<input name="&nome_pesos_variaveis(i)&" type=""hidden"" id="&nome_pesos_variaveis(i)&" class=""peso"" value="&valor_peso&">"	
'			else	
'				linha_pesos(i)="<input name="&nome_pesos_variaveis(i)&" type="&tipo&" id="&nome_pesos_variaveis(i)&" class=""peso"" value="&valor_peso&">"
'			end if	
'			
'		end if				
' %>
<!--    <td width="<%response.Write(width)%>" class="<%response.Write(classe_peso)%>"><div align="<%response.Write(align)%>"><%response.Write(linha_pesos(i))%></div></td>-->
<%'	next%>
</tr>
<%'end if%>
  <tr> 
 <% for j= 0 to ubound(linha_nome_colunas)
 		if j=0 then
			width=width_nom_disciplina
			align="center"
'		elseif j=1 then	
'			width=width_nom_aluno
'			align="left"			
		else
			width=width_else
			align="center"
		end if							
 %>
    <td width="<%response.Write(width)%>" class="<%response.Write(classe_subtit)%>"><div align="<%response.Write(align)%>"><%response.Write(linha_nome_colunas(j))%></div></td>
<%	next%>  
  </tr>
  <%
check = 2

For coe=0 to ubound(co_materia_exibe)

	co_materia = co_materia_exibe(coe)

	
	if subopcao="imp" Then
		classe = "tabela"
		classe_td_imp= " class = 'tabela'"
	elseif nu_matricula = calc then
		classe = "tb_fundo_linha_erro"
		onblur="mudar_cor_blur_erro"	
		classe_td_imp= ""	  	   
	else
		if check mod 2 =0 then
			classe = "tb_fundo_linha_par" 
			onblur="mudar_cor_blur_par"
		else 
			classe ="tb_fundo_linha_impar"
			onblur="mudar_cor_blur_impar"
		end if 
		classe_td_imp= ""		
	end if


	if co_materia="MED" then
		if subopcao="imp" Then
			classe = "tabela"
		else
			classe="tb_fundo_linha_falta"
		end if	
		no_materia="&nbsp;&nbsp;&nbsp;M&eacute;dia"	
		id_tr=""	
    else
	
		Set RS1a = Server.CreateObject("ADODB.Recordset")
		SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&co_materia&"'"
		RS1a.Open SQL1a, CON0
			
		if RS1a.eof then
			response.Write("Erro - Busca pela disciplina")
			response.End()
		else
			co_materia_principal=RS1a("CO_Materia_Principal")	
			if isnull(co_materia_principal)	or co_materia_principal = "" then
				co_materia_principal = co_materia
			end if			
		end if		
		
		Set RSprog = Server.CreateObject("ADODB.Recordset")
		SQLprog = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia = '"& co_materia &"'"
		RSprog.Open SQLprog, CON0

		mae=RSprog("IN_MAE")
		fil=RSprog("IN_FIL")
		in_co=RSprog("IN_CO")
		nu_peso=RSprog("NU_Peso")
		ordem=RSprog("NU_Ordem_Boletim")
		
		if mae=TRUE and fil = TRUE then
			edita="N"
			id_tr=""	
		else	
			linha_tabela=linha_tabela+1 
			edita="S"
			id_tr="celula"&linha_tabela
		end if		
		Set RS1a = Server.CreateObject("ADODB.Recordset")
		SQL1a = "SELECT * FROM TB_Materia WHERE CO_Materia='"&co_materia&"'"
		RS1a.Open SQL1a, CON0
			
		no_materia=RS1a("NO_Materia")
		
		if mae = FALSE then
			no_materia="&nbsp;&nbsp;&nbsp;"&no_materia
		end if
					
	end if			

	%>   
    <tr class="<%response.Write(classe)%>" id="<%response.Write(id_tr)%>">
    <td width="<%response.Write(width_nom_disciplina)%>" <%response.Write(classe_td_imp)%>>
    <%response.Write(no_materia)%>    
    </td>         
     <% 
    Set RS3 = Server.CreateObject("ADODB.Recordset")
    SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& co_materia_principal &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
    Set RS3 = CON_N.Execute(SQL_N)			 
    coluna=0	 
     for n= 0 to ubound(nome_variaveis)
        width=width_else
        align="center"
        'nome_campo=co_materia_principal&"_"&co_materia&"_"&nome_variaveis(n)
    	nome_campo=co_materia_principal&"_"&co_materia&"_"&nome_variaveis(n)&"_1"
        if RS3.EOF then 
            valor=""
        else
            if opt="err6" and nu_matricula = calc then	
                if errou=nome_variaveis(n) then
                    valor=qerrou	
                else
                    valor=Session(nome_variaveis(n))
                end if
            else	
                if tb = "TB_NOTA_V" and (variaveis_bd(n) = "VA_Media1" or variaveis_bd(n) = "VA_Bonus" or variaveis_bd(n) = "VA_Media2" or variaveis_bd(n) = "VA_Rec" or variaveis_bd(n) = "VA_Media3") then
                    Set RS3F = Server.CreateObject("ADODB.Recordset")
                    SQL_NF = "Select * from TB_NOTA_F WHERE CO_Matricula = "& nu_matricula & " AND CO_Materia_Principal = '"& mat_princ &"' AND CO_Materia = '"& co_materia &"' AND NU_Periodo="&periodo
                    Set RS3F = CON_NF.Execute(SQL_NF)	
                    
                    valor=RS3F(""&variaveis_bd(n)&"")															
                else
                    valor=RS3(""&variaveis_bd(n)&"")
                end if	
            end if							
        end if
        
        if (valor="" or isnull(valor)) and subopcao="imp" then
            coluna=coluna+1	
            conteudo="&nbsp;"			
        else
			if nome_variaveis(n) = "rs" or nome_variaveis(n)="rb" then
				tipo_form = "checkbox"
				status_form = "disabled"
			else
				tipo_form = tipo
				status_form = ""				
			end if
		
            if co_materia="MED" or edita="N" or nome_variaveis(n)="media_teste" or nome_variaveis(n)="media_prova" or nome_variaveis(n)="media1" or nome_variaveis(n)="media2" or nome_variaveis(n)="media3" or nome_variaveis(n)="alterado" or nome_variaveis(n)="data_altera" then
                coluna=coluna	
                conteudo=valor
             else
                coluna=coluna+1
                if comunica="s" or subopcao="blq" then
    '						if comunica="s" or (periodo_bloqueado="s" and sistema_origem="WN") then
    
                    conteudo=valor&"<input name='"&nome_campo&"' type='"&tipo_form&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""mudar_cor_focus(celula"&linha_tabela&");javascript:this.select();"" onBlur="&onblur&"(celula"&linha_tabela&") value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_disciplinas&")"" "&status_form&">"		
                else
                    conteudo="<input name='"&nome_campo&"' type='"&tipo_form&"' id='"&linha_tabela&"c"&coluna&"' onFocus=""mudar_cor_focus(celula"&linha_tabela&");javascript:this.select();"" onBlur="&onblur&"(celula"&linha_tabela&") value='"&valor&"' class=""nota"" size=""4"" maxlength=""3"" onkeydown=""keyPressed(this.id,event,"&notas_a_lancar&","&qtd_disciplinas&")"" "&status_form&">"										
                end if
            end if	
        end if	
        'conteudo=n
    %>
    <td width="<%response.Write(width)%>" <%response.Write(classe_td_imp)%>>
        <div align="<%response.Write(align)%>">
            <%response.Write(conteudo)%> 
        </div>
     </td>
    <%	next  
    %>
    </tr>
    <%			
	
	check = check+1 
Next

if subopcao="imp" then
else
%>
    <tr> 
      <td colspan="<%response.Write(qtd_colunas)%>" class="tb_subtit_lanca_notas">
     <%	  
	if funcao_origem="EPN" or subopcao="blq" then
	%>
				<table width="100%" border="0" cellspacing="0">
		  <tr>
			<td colspan="3">
			  <hr>
			</td>
		  </tr>	 
			<tr> 
			<td width="33%">
			<div align="center">
				<input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','index.asp?nvg=<%response.Write(chave)%>');return document.MM_returnValue" value="Voltar">
			  </div>
			  </td>
			  <td width="34%"> <div align="center">
				</div></td>
			  <td width="33%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
	
				  </font></div></td>
			</tr>
		  </table>
	<%
'	elseif  periodo_bloqueado="s" and sistema_origem="WN" then
'	 %>
<!--			<table width="100%" border="0" cellspacing="0">
		  <tr>
			<td colspan="3">
			  <hr>
			</td>
		  </tr>	 
			<tr> 
			<td width="33%">
			<div align="center">
				<input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','altera.asp');return document.MM_returnValue" value="Voltar">
			  </div>
			  </td>
			  <td width="34%"> <div align="center">
				</div></td>
			  <td width="33%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
	
				  </font></div></td>
			</tr>
		  </table>
-->	<%elseif comunica="s" then%>
		 <table width="100%" border="0" cellspacing="0">
		  <tr>
			<td colspan="3">
			  <hr>
			</td>
		  </tr>	 
			<tr> 
			<td width="33%"><div align="center">
							<input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','index.asp?nvg=<%response.Write(chave)%>');return document.MM_returnValue" value="Voltar">
						  </div></td>
			  <td width="34%"> <div align="center"> 
				  <!--<input name="Button" type="button" class="botao_cancelar" onClick="MM_goToURL('parent','notas.asp?ori=02&amp;opt=ok&amp;obr=<%=obr%>');return document.MM_returnValue" value="Cancelar">-->
				</div></td>
			  <td width="33%"> <div align="center"><font size="3" face="Courier New, Courier, mono"> 
				  <input type="submit" name="Submit" value="Confirmar" class="botao_prosseguir">
				  <input name="unidade" type="hidden" id="unidade" value="<%=unidades%>">
				  <input name="curso" type="hidden" id="curso" value="<%=grau%>">
				  <input name="etapa" type="hidden" id="etapa" value="<%=serie%>">
				  <input name="turma" type="hidden" id="turma" value="<%=turma%>">
				  <input name="co_materia" type="hidden" id="co_materia" value="<%= co_materia%>">
				  <input name="periodo" type="hidden" id="periodo" value="<%= periodo%>">
				  <input name="co_prof" type="hidden" id="co_prof" value="<% = co_prof%>">
				  <input name="max" type="hidden" id="max" value="1">
				  <input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
				  <input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
				  </font></div></td>
			</tr>
		  </table>
	<%else%>            
		  <table width="100%" border="0" align="center" cellspacing="0">
			  <tr> 
				<td colspan="3"><hr></td>
			  </tr>
			  <tr> 
				<td width="33%"><div align="center">             
					<input name="bt" type="button" class="botao_cancelar" id="bt" onClick="MM_goToURL('parent','index.asp?nvg=<%response.Write(chave)%>');return document.MM_returnValue" value="Voltar">
				  </div></td>
				<td width="34%"><div align="center"> 
					<!--<input name="Submit" type="button" class="botao_prosseguir_comunicar" onClick="MM_goToURL('parent','notas.asp?or=01&opt=cln&obr=<%=obr%>');return document.MM_returnValue" value="Comunicar ao Coordenador T&eacute;rmino da Planilha">-->&nbsp;
				  </div></td>
				<td width="33%"> <div align="center"> 
					<input type="submit" name="Submit2" value="Salvar" class="botao_prosseguir">
					<input name="unidade" type="hidden" id="unidade" value="<%=unidade%>">
					<input name="curso" type="hidden" id="curso" value="<%=curso%>">
					<input name="etapa" type="hidden" id="etapa" value="<%=etapa%>">
					<input name="turma" type="hidden" id="turma" value="<%=turma%>">
					<input name="endereco_origem" type="hidden" id="endereco_origem" value="<%= endereco_origem%>">
					<input name="periodo" type="hidden" id="periodo" value="<%= periodo%>">
					<input name="nu_matricula" type="hidden" id="co_prof" value="<% = nu_matricula%>">
					<input name="max" type="hidden" id="max" value="1">
					<input name="co_usr" type="hidden" id="co_usr" value="<% = co_usr%>">
					<input name="ano_letivo" type="hidden" id="ano_letivo" value="<% = ano_letivo%>">
				  </div></td>
			  </tr>
			</table>
	<%end if
end if%>        

</form>    

