<%On Error Resume Next%>
<!--#include file="caminhos.asp"-->
<!--#include file="funcoes6.asp"-->
<!--#include file="../../global/tabelas_escolas.asp"-->
<%
opt=request.querystring("opt")
nome = session("nome") 
acesso = session("acesso")
co_usr = session("co_user")
permissao = session("permissao") 
ano_letivo = session("ano_letivo") 
sistema_local=session("sistema_local")
grupo=session("grupo")
escola=session("escola")
chave=session("chave")
escola= session("escola")
		
this_file = Request.ServerVariables("SCRIPT_NAME")
arPath = Split(this_Path, "/")

Function URLDecode(sConvert)
    Dim aSplit
    Dim sOutput
    Dim I
    If IsNull(sConvert) Then
       URLDecode = ""
       Exit Function
    End If

    ' convert all pluses to spaces
    sOutput = REPLACE(sConvert, "+", " ")

    ' next convert %hexdigits to the character
    aSplit = Split(sOutput, "%")

    If IsArray(aSplit) Then
      sOutput = aSplit(0)
      For I = 0 to UBound(aSplit) - 1
        sOutput = sOutput & _
          Chr("&H" & Left(aSplit(i + 1), 2)) &_
          Right(aSplit(i + 1), Len(aSplit(i + 1)) - 2)
      Next
    End If

    URLDecode = sOutput
End Function

nivel=4

	Set CON_AL = Server.CreateObject("ADODB.Connection") 
	ABRIR_AL = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_AL.Open ABRIR_AL
	
	Set CON0 = Server.CreateObject("ADODB.Connection") 
	ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON0.Open ABRIR0
	
	Set CON3 = Server.CreateObject("ADODB.Connection") 
	ABRIR3 = "DBQ="& CAMINHO_o & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON3.Open ABRIR3	
	
	Set CONCONT = Server.CreateObject("ADODB.Connection") 
	ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONCONT.Open ABRIRCONT	
	
	Set CON_WF = Server.CreateObject("ADODB.Connection") 
	ABRIR_WF= "DBQ="& CAMINHO_wf & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_WF.Open ABRIR_WF	
	
	Set CON_wr = Server.CreateObject("ADODB.Connection") 
	ABRIR_wr = "DBQ="& CAMINHO_wr & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON_wr.Open ABRIR_wr		

	Set CONg = Server.CreateObject("ADODB.Connection") 
	ABRIRg = "DBQ="& CAMINHO_g & ";Driver={Microsoft Access Driver (*.mdb)}"
	CONg.Open ABRIRg
	
	Set CON7 = Server.CreateObject("ADODB.Connection") 
	ABRIR7 = "DBQ="& CAMINHO_h & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON7.Open ABRIR7		

if nome = "" or acesso = "" or co_usr = "" or permissao = "" or ano_letivo = "" or chave = "" or isnull(chave) then
	if nivel=0 then
	response.Redirect("default.asp?opt=00")
	elseif nivel=1 then
	response.Redirect("../default.asp?opt=00")
	elseif nivel=2 then
	response.Redirect("../../default.asp?opt=00")
	elseif nivel=3 then
	response.Redirect("../../../default.asp?opt=00")
	elseif nivel=4 then
	response.Redirect("../../../../default.asp?opt=00")
	end if
end if
		
%>
<%if opt="esq" then
modelo = ucase(Request.Form("e_pub"))
modelo = replace(modelo,"SUPASUP", "ª")
modelo = replace(modelo,"SUPOSUP", "º")
modelo = replace(modelo,"23A23", "Á")
modelo = replace(modelo,"23E23", "É")
modelo = replace(modelo,"23I23", "Í")
modelo = replace(modelo,"23O23", "Ó")
modelo = replace(modelo,"23U23", "Ú")
modelo = replace(modelo,"23C23", "Ç")
modelo = replace(modelo,"45A45", "Ã")
modelo = replace(modelo,"45N45", "Ñ")
modelo = replace(modelo,"45O45", "Õ")
modelo = replace(modelo,"78A78", "Â")
modelo = replace(modelo,"78E78", "Ê")
modelo = replace(modelo,"78O78", "Ô")
esquema_executa=modelo

%>
<form id="form1" name="form1" method="post" action="">
<table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td colspan="5"><table id="tblInnerHTML" width="100%" border="0" cellspacing="0" cellpadding="0">
  <%
 
 		Set RSTESQ = Server.CreateObject("ADODB.Recordset")
		SQLTESQ = "SELECT count(NO_Materia) as Qtd_Materias FROM TB_Historico_Esquema_Disciplinas where UCASE(NO_Esquema) = '"&esquema_executa&"'"
		RSTESQ.Open SQLTESQ, CON7  
		
		qtd_materias = RSTESQ("Qtd_Materias")  
  
		Set RSESQ = Server.CreateObject("ADODB.Recordset")
		SQLESQ = "SELECT NO_Materia FROM TB_Historico_Esquema_Disciplinas where UCASE(NO_Esquema) = '"&esquema_executa&"' order by NO_Materia"
		RSESQ.Open SQLESQ, CON7  
  registro = 1
  while not RSESQ.EOF   
	materia_esquema = RSESQ("NO_Materia")
	if registro<qtd_materias then
		img_bt="close.png"  	
		'acao="deleteRow("&registro&");"
		alt="Excluir Disciplina"	
		classe = "remove"	
	else 
		img_bt="add.png"  
		'acao="addRow();changeImage("&registro&")"
		alt="Adicionar Disciplina"
		classe = "add"			
	end if	
%> 
            <tr>
              <td width="80" align="right" class="form_corpo">
          <input name="num_linha" type="hidden" id="num_linha" value="<%response.Write(registro)%>">Disciplina:&nbsp;</td>
              <td width="380" class="form_corpo"><span class="ui-widget">
                <input class="textInput" name="disciplina_<%response.Write(registro)%>" type="text" id="disciplina_<%response.Write(registro)%>" size="50" value="<%response.Write(Server.URLEncode(materia_esquema))%>" />
              </span></td>
              <td width="100" align="right" class="form_corpo"><%response.Write(Server.URLEncode("Carga Hor&aacute;ria:&nbsp;"))%></td>
              <td width="40" class="form_corpo"><input class="textInput" name="carga_form_<%response.Write(registro)%>" type="text" id="carga_form_<%response.Write(registro)%>" size="6" maxlength="4" /></td>
              <td width="80" align="right" class="form_corpo"><%response.Write(Server.URLEncode("Frequ&ecirc;ncia:&nbsp;"))%></td>
              <td width="40" class="form_corpo"><input name="frequencia_form_<%response.Write(registro)%>" type="text" class="textInput" id="frequencia_form_<%response.Write(registro)%>" size="6" maxlength="4" /></td>
              <td width="50" align="right" class="form_corpo">Nota:&nbsp;</td>
              <td width="40" class="form_corpo"><input name="nota_form_<%response.Write(registro)%>" type="text" class="textInput" id="nota_form_<%response.Write(registro)%>" size="6" maxlength="4" /></td>
              <td width="80" align="right" class="form_corpo">Aprovado:&nbsp;</td>
              <td width="60" class="form_corpo"><select name="aprovado_<%response.Write(registro)%>" class="select_style">
                <option value="S" selected="selected">Sim</option>
                <option value="N">
                  <%response.Write(Server.URLEncode("N&atilde;o"))%>
                  </option>
              </select></td>
              <td width="50" class="form_corpo"><div id="<%response.Write(registro)%>"><a href="javascript:void(0);"  class="<%response.Write(classe)%>"><img src="../../../../img/<%response.Write(img_bt)%>" width="20" height="20"  alt="<%response.Write(alt)%>" border="0" /></a></div></td>
            </tr>
          
<%
registro = registro+1
RSESQ.MOVENEXT
WEND%> 
</table></td>
          </tr>  
    <tr>
      <td colspan="5"><hr /></td>
    </tr>
        <tr>
          <td width="20%">&nbsp;<input name="qtd_itens" type="hidden" id="qtd_itens" value="<%response.Write(qtd_materias)%>">
          <input name="itens_criados" type="hidden" id="itens_criados" value="<%response.Write(qtd_materias)%>"></td>
          <td width="20%" align="center"><input type="button" name="button" id="button" class="botao_cancelar" value="Cancelar" /></td>
          <td width="20%">&nbsp;</td>
          <td width="20%" align="center"><input type="button" name="button" id="button" class="botao_prosseguir" value="Salvar" /></td>
          <td width="20%">&nbsp;</td>
        </tr>
      </table></td>
    </tr>
  </table>
</form>
<%elseif opt="gesq" then
nomeDoEsquema=Request.Form("n_pub")
modeloDoEsquema=Request.Form("m_pub")

	Set RS0 = Server.CreateObject("ADODB.Recordset")
	SQL0 = "SELECT * FROM TB_Historico_Esquema_Disciplinas where NO_Esquema='"&nomeDoEsquema&"'"
	RS0.Open SQL0, CON7

	if RS0.EOF then
		wrk_apaga="N"
	else
		wrk_apaga="S"	
	end if

	IF wrk_apaga="S" then
		Set RSD = Server.CreateObject("ADODB.Recordset")
		SQLD = "DELETE  * from TB_Historico_Esquema_Disciplinas where NO_Esquema='"&nomeDoEsquema&"'"
		RSD.Open SQLD, CON7	
	end if
	
	if modeloDoEsquema<> "nulo" then
		modeloDoEsquema = URLDecode(modeloDoEsquema)
		disciplinasModelo = split(modeloDoEsquema,",")
		
		for dm=0 to ubound(disciplinasModelo)
			Set RS = server.createobject("adodb.recordset")		
			RS.open "TB_Historico_Esquema_Disciplinas", CON7, 2, 2 'which table do you want open
			RS.addnew
			
				RS("NO_Esquema") = nomeDoEsquema
				RS("NO_Materia") = disciplinasModelo(dm)
						
			RS.update
			set RS=nothing		
		next
		end if
%>
<select name="carrega_modelo" class="select_style" id="carrega_modelo" onChange="javascript:verifica_habilitacao(this.id,this.value,'bt_carrega_esquema');javascript:verifica_habilitacao(this.id,this.value,'bt_exclui_esquema');">
                        <option value="nulo"></option>
                        <%		
		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT NO_Esquema FROM TB_Historico_Esquema_Disciplinas group by NO_Esquema order by NO_Esquema"
		RS0.Open SQL0, CON7
		
While not RS0.EOF
no_esquema = RS0("NO_Esquema")

if no_esquema = nomeDoEsquema then
	esquemsSelected = "selected"
else
	esquemsSelected = ""
end if	

modelo = ucase(no_esquema)
modelo = replace(modelo, "ª","SUPASUP")
modelo = replace(modelo, "º","SUPOSUP")
modelo = replace(modelo, "Á","23A23")
modelo = replace(modelo, "É","23E23")
modelo = replace(modelo, "Í","23I23")
modelo = replace(modelo, "Ó","23O23")
modelo = replace(modelo, "Ú","23U23")
modelo = replace(modelo, "Ç","23C23")
modelo = replace(modelo, "Ã","45A45")
modelo = replace(modelo, "Ñ","45N45")
modelo = replace(modelo, "Õ","45O45")
modelo = replace(modelo, "Â","78A78")
modelo = replace(modelo, "Ê","78E78")
modelo = replace(modelo, "Ô","78O78")
%>
                        <option value="<%response.Write(modelo)%>" <%response.Write(esquemsSelected)%>>
                          <%response.Write(Server.URLEncode(no_esquema))%>
                        </option>
                        <%RS0.MOVENEXT
WEND
%>
                      </select>&nbsp;&nbsp;
                      <input name="bt_carrega_esquema" type="submit" class="botao_prosseguir" id="bt_carrega_esquema" value="Carregar" disabled onClick="carregaEsquema(carrega_modelo.value)">
                      <input name="bt_exclui_esquema" type="submit" class="botao_excluir" id="bt_exclui_esquema" value="Excluir" disabled onClick="excluiEsquema(carrega_modelo.value,'nulo')">
<%
elseif opt="s" then
tp_curso_altera=Request.Form("t_pub")
%>
<select name="co_seg" class="select_style" id="co_seg">
   <option value="nulo" selected></option>
<%
if tp_curso_altera<>"nulo" then

	Set RS0 = Server.CreateObject("ADODB.Recordset")
	SQL0 = "SELECT * FROM TB_Segmento where TP_Curso='"&tp_curso_altera&"' order by NU_Ordem"
	RS0.Open SQL0, CON7


	While not RS0.EOF
	
	co_seqmento = RS0("CO_Seg")		
	no_seqmento = RS0("NO_Abreviado_Curso")		
	%>
							<option value="<%response.Write(co_seqmento)%>"> 
							<%response.Write(Server.URLEncode(no_seqmento))%>
							</option>
							<%
	RS0.MOVENEXT
	WEND
end if	
%>
</select>
<%elseif opt="c" then
unidade_altera=Request.Form("u_pub")
if unidade_altera="" or isnull(unidade_altera) then
session("u_pub")=session("u_pub")
else
session("u_pub")=unidade_altera
session("u_aoc")=unidade_altera
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if
%>
                      <select name="curso" class="select_style" id="curso" onchange="recuperarEtapa(this.value)">
                        <option value="999990" selected></option>
                        <%		

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT DISTINCT CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")
		RS0.Open SQL0, CON0
		
		
While not RS0.EOF
CO_Curso = RS0("CO_Curso")

		Set RS0a = Server.CreateObject("ADODB.Recordset")
		SQL0a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
		RS0a.Open SQL0a, CON0
		
NO_Curso = RS0a("NO_Abreviado_Curso")		
%>
                        <option value="<%response.Write(CO_Curso)%>"> 
                        <%response.Write(Server.URLEncode(NO_Curso))%>
                        </option>
                        <%
RS0.MOVENEXT
WEND
%>
</select>
<%elseif opt="c2" then
unidade_altera=Request.Form("u_pub")
if unidade_altera="" or isnull(unidade_altera) then
session("u_pub")=session("u_pub")
else
session("u_pub")=unidade_altera
session("u_aoc")=unidade_altera
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if
%>
                      <select name="curso" class="select_style" id="curso">
                        <option value="999990" selected></option>
                        <%		

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT DISTINCT CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")
		RS0.Open SQL0, CON0
		
		
While not RS0.EOF
CO_Curso = RS0("CO_Curso")

		Set RS0a = Server.CreateObject("ADODB.Recordset")
		SQL0a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
		RS0a.Open SQL0a, CON0
		
NO_Curso = RS0a("NO_Abreviado_Curso")		
%>
                        <option value="<%response.Write(CO_Curso)%>"> 
                        <%response.Write(Server.URLEncode(NO_Curso))%>
                        </option>
                        <%
RS0.MOVENEXT
WEND
%>
</select>
<%elseif opt="c3" then
unidade_altera=Request.Form("u_pub")
if unidade_altera="" or isnull(unidade_altera) then
session("u_pub")=session("u_pub")
else
session("u_pub")=unidade_altera
session("u_aoc")=unidade_altera
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if
%>
                      <select name="curso" class="select_style" id="curso" onchange="recuperarControle(this.value)">
                        <option value="999990" selected></option>
                        <%		

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT DISTINCT CO_Curso FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")
		RS0.Open SQL0, CON0
		
		
While not RS0.EOF
CO_Curso = RS0("CO_Curso")

		Set RS0a = Server.CreateObject("ADODB.Recordset")
		SQL0a = "SELECT * FROM TB_Curso where CO_Curso='"&CO_Curso&"'"
		RS0a.Open SQL0a, CON0
		
NO_Curso = RS0a("NO_Abreviado_Curso")		
%>
                        <option value="<%response.Write(CO_Curso)%>"> 
                        <%response.Write(Server.URLEncode(NO_Curso))%>
                        </option>
                        <%
RS0.MOVENEXT
WEND
%>
</select>
<%elseif opt="e" then

curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("c_aoc")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="etapa" class="select_style" id="etapa" onchange="recuperarTurma(this.value)">
                        <option value="999990" selected></option>
                        <%		
		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"'"
		RS0b.Open SQL0b, CON0
				
While not RS0b.EOF
CO_Etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub")&"' AND CO_Etapa='"&CO_Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
%>
                        <option value="<%response.Write(CO_Etapa)%>"> 
                        <%response.Write(Server.URLEncode(NO_Etapa))%>
                        </option>
                        <%
RS0b.MOVENEXT
WEND
%>
                      </select>
<%
'Essa combo da etapa chama também a rotina da disciplina
elseif opt="e2" then

curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="etapa" class="select_style" id="etapa" onchange="recuperarTurma(this.value);recuperarDisciplina(this.value)">
                        <option value="999990" selected></option>
                        <%		
		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"'"
		RS0b.Open SQL0b, CON0
				
While not RS0b.EOF
CO_Etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub")&"' AND CO_Etapa='"&CO_Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
%>
                        <option value="<%response.Write(CO_Etapa)%>"> 
                        <%response.Write(Server.URLEncode(NO_Etapa))%>
                        </option>
                        <%
RS0b.MOVENEXT
WEND
%>
                      </select>
<%'Essa combo da etapa chama também a rotina da avaliação
elseif opt="e3" then

curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="etapa" class="select_style" id="etapa" onchange="recuperarPeriodo(this.value);recuperarAvaliacoes(this.value)">
                        <option value="999990" selected></option>
                        <%		
		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"'"
		RS0b.Open SQL0b, CON0
				
While not RS0b.EOF
CO_Etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub")&"' AND CO_Etapa='"&CO_Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
%>
                        <option value="<%response.Write(CO_Etapa)%>"> 
                        <%response.Write(Server.URLEncode(NO_Etapa))%>
                        </option>
                        <%
RS0b.MOVENEXT
WEND
%>
                      </select>
<%'Essa combo da etapa chama também a rotina do Período
elseif opt="e4" then

curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="etapa" class="select_style" id="etapa" onchange="recuperarPeriodo()">
                        <option value="999990" selected></option>
                        <%		

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"'"
		RS0b.Open SQL0b, CON0
				
While not RS0b.EOF
CO_Etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub")&"' AND CO_Etapa='"&CO_Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
%>
                        <option value="<%response.Write(CO_Etapa)%>"> 
                        <%response.Write(Server.URLEncode(NO_Etapa))%>
                        </option>
                        <%
RS0b.MOVENEXT
WEND
%>
</select>
<%
'Essa combo da etapa chama também a rotina da disciplina
elseif opt="e5" then

curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="etapa" class="select_style" id="etapa" onchange="recuperarDisciplina(curso.value,this.value)">
                        <option value="999990" selected></option>
                        <%		
		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"'"
		RS0b.Open SQL0b, CON0
				
While not RS0b.EOF
CO_Etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub")&"' AND CO_Etapa='"&CO_Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
%>
                        <option value="<%response.Write(CO_Etapa)%>"> 
                        <%response.Write(Server.URLEncode(NO_Etapa))%>
                        </option>
                        <%
RS0b.MOVENEXT
WEND
%>
                      </select>					  					  					  
<%
'Essa combo da etapa chama também a rotina da disciplina
elseif opt="e6" then

curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="etapa" class="select_style" id="etapa" onchange="MM_callJS('submitfuncao()')">
                        <option value="999990" selected></option>
                        <%		
		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"'"
		RS0b.Open SQL0b, CON0
				
While not RS0b.EOF
CO_Etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub")&"' AND CO_Etapa='"&CO_Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
%>
                        <option value="<%response.Write(CO_Etapa)%>"> 
                        <%response.Write(Server.URLEncode(NO_Etapa))%>
                        </option>
                        <%
RS0b.MOVENEXT
WEND
%>

                      </select>	 
<%elseif opt="e7" then

curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("c_aoc")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="etapa" class="select_style" id="etapa" onchange="recuperarControle(this.value)">
                        <option value="999990" selected></option>
                        <%		
		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"'"
		RS0b.Open SQL0b, CON0
				
While not RS0b.EOF
CO_Etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub")&"' AND CO_Etapa='"&CO_Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
%>
                        <option value="<%response.Write(CO_Etapa)%>"> 
                        <%response.Write(Server.URLEncode(NO_Etapa))%>
                        </option>
                        <%
RS0b.MOVENEXT
WEND
%>
                      </select> 
<%
elseif opt="e8" then

curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="etapa" id="etapa" class="select_style" onchange="recuperarTurma(this.value),recuperarPeriodo(this.value)">
                        <option value="999990" selected></option>
                        <%		

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"'"
		RS0b.Open SQL0b, CON0
				
While not RS0b.EOF
CO_Etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub")&"' AND CO_Etapa='"&CO_Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
%>
                        <option value="<%response.Write(CO_Etapa)%>"> 
                        <%response.Write(Server.URLEncode(NO_Etapa))%>
                        </option>
                        <%
RS0b.MOVENEXT
WEND
%>
</select>  
<%elseif opt="e9" then
'recupera etapas maiores que 6 quando curso = EF
curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("c_aoc")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if
curso_altera=curso_altera*1
if curso_altera=1 then
	sql_adic = " AND CO_Etapa>'5'"
else
	sql_adic = ""
end if

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"'"&sql_adic
		RS0b.Open SQL0b, CON0

%>
                      <select name="etapa" class="select_style" id="etapa" onchange="recuperarTurma(this.value)">
                        <option value="999990" selected></option>
                        <%		

				
While not RS0b.EOF
CO_Etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub")&"' AND CO_Etapa='"&CO_Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
%>
                        <option value="<%response.Write(CO_Etapa)%>"> 
                        <%response.Write(Server.URLEncode(NO_Etapa))%>
                        </option>
                        <%
RS0b.MOVENEXT
WEND
%>
                      </select>    
<%elseif opt="e10" then
'recupera etapas maiores que 6 quando curso = EF
curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("c_aoc")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if


		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"'"
		RS0b.Open SQL0b, CON0

%>
                      <select name="etapa" class="select_style"  id="etapa">
                        <option value="999990" selected></option>
                        <%		

				
While not RS0b.EOF
CO_Etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub")&"' AND CO_Etapa='"&CO_Etapa&"'"
		RS0c.Open SQL0c, CON0
		
NO_Etapa = RS0c("NO_Etapa")		
%>
                        <option value="<%response.Write(CO_Etapa)%>"> 
                        <%response.Write(Server.URLEncode(NO_Etapa))%>
                        </option>
                        <%
RS0b.MOVENEXT
WEND
%>
                      </select>                                                              
<%elseif opt="t" then

etapa_altera=Request.Form("e_pub")
if etapa_altera="" or isnull(etapa_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=etapa_altera
session("t_pub")=session("t_pub")
end if
%>
<select name="turma" class="select_style"  id="turma">
                        <option value="999990" selected></option>
    <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"' AND CO_Etapa='" & session("e_pub") & "' order by CO_Turma" 
		RS3.Open SQL3, CON0						

while not RS3.EOF
co_turma= RS3("CO_Turma")
 %>
<option value="<%=co_turma%>"> 
    <%response.Write(Server.URLEncode(co_turma))%>
</option>
    <%
RS3.MOVENEXT
WEND
%>
    
</select>
 <%
'Essa combo da turma chama também a submete o form 
 elseif opt="t2" then

etapa_altera=Request.Form("e_pub")
if etapa_altera="" or isnull(etapa_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=etapa_altera
session("t_pub")=session("t_pub")
end if
%>
<select name="turma" class="select_style" id="turma" onChange="MM_callJS('submitfuncao()')">
                        <option value="999990" selected></option>
    <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"' AND CO_Etapa='" & session("e_pub") & "' order by CO_Turma" 
		RS3.Open SQL3, CON0						

while not RS3.EOF
co_turma= RS3("CO_Turma")

 %>
<option value="<%=co_turma%>"> 
    <%response.Write(Server.URLEncode(co_turma))%>
</option>
    <%
RS3.MOVENEXT
WEND
%>
    
</select>
  <%
'Essa combo da turma chama também a submete o form 
 elseif opt="t3" then

etapa_altera=Request.Form("e_pub")
if etapa_altera="" or isnull(etapa_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=etapa_altera
session("t_pub")=session("t_pub")
end if
%>
<select name="turma" class="select_style" id="turma" onChange="MM_callJS('recuperarDisciplina()')">
                        <option value="999990" selected></option>
    <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"' AND CO_Etapa='" & session("e_pub") & "' order by CO_Turma" 
		RS3.Open SQL3, CON0						

while not RS3.EOF
co_turma= RS3("CO_Turma")

 %>
<option value="<%=co_turma%>"> 
    <%response.Write(Server.URLEncode(co_turma))%>
</option>
    <%
RS3.MOVENEXT
WEND
%>
    
</select>
  <%elseif opt="t4" then

etapa_altera=Request.Form("e_pub")
if etapa_altera="" or isnull(etapa_altera) then
session("c_pub")=session("c_pub")
else
session("e_pub")=etapa_altera
session("t_pub")=session("t_pub")
end if
%>
 <select name="turma" class="select_style" id="turma" onchange="recuperarPeriodo()">
                        <option value="999990" selected></option> 
    <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Turma where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"' AND CO_Etapa='" & session("e_pub") & "' order by CO_Turma" 
		RS3.Open SQL3, CON0
		
							
co_turma_check=9999990
while not RS3.EOF
co_turma= RS3("CO_Turma")

if co_turma = co_turma_check then
RS3.MOVENEXT
else


 %>
<option value="<%=co_turma%>"> 
    <%response.Write(Server.URLEncode(co_turma))%>
</option>
    <%
co_turma_check = co_turma
RS3.MOVENEXT
end if
WEND
%>
    
 </select>
  <%

' Só serve para gravar o valor de turma para as funções atualizar ocorrências (AOC), Emitir Ficha e Emitir Boletim do Web Secretaria
elseif opt="t5" then

etapa_altera=Request.Form("e_pub")
if etapa_altera="" or isnull(etapa_altera) then
session("c_pub")=session("c_pub")
else
session("e_pub")=etapa_altera
session("e_aoc")=etapa_altera
session("t_pub")=session("t_pub")
end if
%>
 <select name="turma" class="select_style" id="turma" onchange="gravarTurma(this.value)">
                        <option value="999990" selected></option> 
    <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Turma where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"' AND CO_Etapa='" & session("e_pub") & "' order by CO_Turma" 
		RS3.Open SQL3, CON0
		
							
co_turma_check=9999990
while not RS3.EOF
co_turma= RS3("CO_Turma")

if co_turma = co_turma_check then
RS3.MOVENEXT
else


 %>
<option value="<%=co_turma%>"> 
    <%response.Write(Server.URLEncode(co_turma))%>
</option>
    <%
co_turma_check = co_turma
RS3.MOVENEXT
end if
WEND
%>
    
 </select> 
 
<%
' Só serve para gravar o valor de turma para a função atualizar ocorrências (AOC)
elseif opt="t6" then

turma_aoc=Request.Form("t_pub")
session("t_aoc")=turma_aoc


'Essa combo da turma chama também a submete o form 
 elseif opt="t7" then

etapa_altera=Request.Form("e_pub")
if etapa_altera="" or isnull(etapa_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=etapa_altera
session("t_pub")=session("t_pub")
end if
%>
<select name="turma" class="select_style" id="turma" onChange="recuperarAvaliacoes(etapa.value);">
                        <option value="999990" selected></option>
    <%
		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT DISTINCT CO_Turma FROM TB_Turma where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"' AND CO_Etapa='" & session("e_pub") & "' order by CO_Turma" 
		RS3.Open SQL3, CON0						

while not RS3.EOF
co_turma= RS3("CO_Turma")

 %>
<option value="<%=co_turma%>"> 
    <%response.Write(Server.URLEncode(co_turma))%>
</option>
    <%
RS3.MOVENEXT
WEND
%>
    
</select>
<%
elseif opt="d" then

curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="mat_prin" class="select_style">
                        <option value="999999" selected></option>
                        <%
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& session("e_pub") &"' AND CO_Curso ='"& session("c_pub") &"' order by NU_Ordem_Boletim "
		RS5.Open SQL5, CON0

while not RS5.EOF
co_mat_prin= RS5("CO_Materia")


		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_prin &"'"
		RS7.Open SQL7, CON0
		
		no_mat_prin= RS7("NO_Materia")
%>
                        <option value="<%=co_mat_prin%>"> 
                        <%response.Write(Server.URLEncode(no_mat_prin))%>						
                        </option>
                        <%

RS5.MOVENEXT
WEND%>
                      </select>
<%
elseif opt="d2" then
prof_altera=Session("co_prof")
curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if


		Set RSG = Server.CreateObject("ADODB.Recordset")
		SQLG = "SELECT DISTINCT CO_Materia_Principal FROM TB_Da_Aula where CO_Professor ="& prof_altera&" and CO_Etapa = '"&session("e_pub") &"' AND NU_Unidade = "&session("u_pub")&" and CO_Curso = '"&session("c_pub") &"' order by CO_Materia_Principal"
		RSG.Open SQLG, CONg
		
IF RSG.EOF THEN

RESPONSE.Write("Sem disciplinas cadastradas. Procure seu Coordenador.")

ELSE
%>
                      <select name="mat_prin" class="select_style" onChange="MM_callJS('recuperarPeriodo()')">
                        <option value="999999" selected></option>
                        <%
while not RSG.EOF
co_mat_prin= RSG("CO_Materia_Principal")

		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_prin &"'"
		RS7.Open SQL7, CON0
		
		no_mat_prin= RS7("NO_Materia")
%>
                        <option value="<%=co_mat_prin%>"> 
                        <%response.Write(Server.URLEncode(no_mat_prin))%>						
                        </option>
                        <%

RSG.MOVENEXT
WEND
END IF
%>
</select>
<%
elseif opt="d3" then
etapa_altera=Request.Form("etapa_pub")
curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=etapa_altera
session("t_pub")=session("t_pub")
end if

%>
                       <select name="mat_prin" class="select_style" onChange="MM_callJS('recuperarPeriodo()')">
                        <option value="999999" selected></option> 
                        <%
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& session("e_pub") &"' AND CO_Curso ='"& session("c_pub") &"' order by NU_Ordem_Boletim "
		RS5.Open SQL5, CON0

'response.Write(SQL5)

while not RS5.EOF
co_mat_prin= RS5("CO_Materia")


		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_mat_prin &"'"
		RS7.Open SQL7, CON0
		
if RS7.eof then
		Set RS7b = Server.CreateObject("ADODB.Recordset")
		SQL7b = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_prin &"'"
		RS7b.Open SQL7b, CON0		
		
		no_mat_prin= RS7b("NO_Materia")
%>
                        <option value="<%=co_mat_prin%>"> 
                        <%response.Write(Server.URLEncode(no_mat_prin))%>						
                        </option>
                        <%
RS5.MOVENEXT						
else
RS5.MOVENEXT
end if
WEND%>

                      </select>					  					  
<%
elseif opt="d4" then

curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")
end if

%>
                      <select name="mat_prin" class="select_style" onChange="MM_callJS('submitfuncao()')">
                        <option value="999999" selected></option>
                        <%
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& session("e_pub") &"' AND CO_Curso ='"& session("c_pub") &"' order by NU_Ordem_Boletim "
		RS5.Open SQL5, CON0

while not RS5.EOF
co_mat_prin= RS5("CO_Materia")


		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_prin &"'"
		RS7.Open SQL7, CON0
		
		no_mat_prin= RS7("NO_Materia")
%>
                        <option value="<%=co_mat_prin%>"> 
                        <%response.Write(Server.URLEncode(no_mat_prin))%>						
                        </option>
                        <%

RS5.MOVENEXT
WEND%>
                      </select>
<%
elseif opt="d5" then
etapa_altera=Request.Form("etapa_pub")
curso_altera=Request.Form("c_pub")
if curso_altera="" or isnull(curso_altera) then
session("c_pub")=session("c_pub")
else
session("u_pub")=session("u_pub")
session("c_pub")=curso_altera
session("e_pub")=etapa_altera
session("t_pub")=session("t_pub")
end if

%>
                       <select name="mat_prin" class="select_style" onChange="MM_callJS('recuperarPeriodo()');MM_callJS('recuperarAvaliacoes()')">
                        <option value="999999" selected></option> 
                        <%
		Set RS5 = Server.CreateObject("ADODB.Recordset")
		SQL5 = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& session("e_pub") &"' AND CO_Curso ='"& session("c_pub") &"' order by NU_Ordem_Boletim "
		RS5.Open SQL5, CON0

'response.Write(SQL5)

while not RS5.EOF
co_mat_prin= RS5("CO_Materia")


		Set RS7 = Server.CreateObject("ADODB.Recordset")
		SQL7 = "SELECT * FROM TB_Materia where CO_Materia_Principal ='"& co_mat_prin &"'"
		RS7.Open SQL7, CON0
		
if RS7.eof then
		Set RS7b = Server.CreateObject("ADODB.Recordset")
		SQL7b = "SELECT * FROM TB_Materia where CO_Materia ='"& co_mat_prin &"'"
		RS7b.Open SQL7b, CON0		
		
		no_mat_prin= RS7b("NO_Materia")
%>
                        <option value="<%=co_mat_prin%>"> 
                        <%response.Write(Server.URLEncode(no_mat_prin))%>						
                        </option>
                        <%
RS5.MOVENEXT						
else
RS5.MOVENEXT
end if
WEND%>

                      </select>							  

<%
elseif opt="p" then
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")

%>
<select name="periodo" class="select_style" id="periodo" onChange="MM_callJS('submitfuncao()')">
                                      <option value="0" selected></option>
                                      <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")%>
                                      <option value="<%=NU_Periodo%>"> 
                                      <%response.Write(Server.URLEncode(NO_Periodo))%>
                                      </option>
                                      <%RS4.MOVENEXT
WEND%>
</select>
<%
elseif opt="p1" then
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")

%>
<select name="periodo" class="select_style" id="periodo">
                                      <option value="0" selected></option>
                                      <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")%>
                                      <option value="<%=NU_Periodo%>"> 
                                      <%response.Write(Server.URLEncode(NO_Periodo))%>
                                      </option>
                                      <%RS4.MOVENEXT
WEND%>
</select>
<%
elseif opt="p2" then
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")

%>
<select name="periodo" class="select_style" id="periodo" onChange="MM_callJS('recuperarAvaliacoes()')">
                                      <option value="0" selected></option>
                                      <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")%>
                                      <option value="<%=NU_Periodo%>"> 
                                      <%response.Write(Server.URLEncode(NO_Periodo))%>
                                      </option>
<%RS4.MOVENEXT
WEND%>
                                    </select>									
<%
elseif opt="p3" then
'etapa_altera=Request.Form("e_pub")
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")

%>
<select name="periodo" class="select_style" id="periodo" onchange="recuperarMedia()">
                                      <option value="0" selected></option>
                                      <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")%>
                                      <option value="<%=NU_Periodo%>"> 
                                      <%response.Write(Server.URLEncode(NO_Periodo))%>
                                      </option>
<%RS4.MOVENEXT
WEND%>
</select>	
<%
'Função ACC
elseif opt="p4" then
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")

%>
<select name="periodo" class="select_style" id="periodo" onchange="GuardaPeriodo(this.value)">
  <option value="0" selected></option>
                                      <%
		Set RS4 = Server.CreateObject("ADODB.Recordset")
		SQL4 = "SELECT * FROM TB_Periodo order by NU_Periodo"
		RS4.Open SQL4, CON0

while not RS4.EOF
NU_Periodo =  RS4("NU_Periodo")
NO_Periodo= RS4("NO_Periodo")%>
                                      <option value="<%=NU_Periodo%>"> 
                                      <%response.Write(Server.URLEncode(NO_Periodo))%>
                                      </option>
                                      <%RS4.MOVENEXT
WEND%>
</select>

<%elseif opt="av" then
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=request.form("e_pub")
session("t_pub")=session("t_pub")


		Set RSTB = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& session("u_pub") &" AND CO_Curso = '"& session("c_pub") &"' AND CO_Etapa = '"& session("e_pub")  &"'"

		Set RSTB = CONg.Execute(CONEXAO)

'response.Write(CONEXAO)
	if RSTB.eof then
	response.Write("<font class=form_dado_texto>n&atilde;o dispon&iacute;vel</font>")
	else
		nota = RSTB("TP_Nota")
		if nota ="TB_NOTA_A" then
		CAMINHOn = CAMINHO_na
		opcao="A"
		else
			if nota="TB_NOTA_B" then
			CAMINHOn = CAMINHO_nb
			opcao="B"
			else
				if nota ="TB_NOTA_C" then
				CAMINHOn = CAMINHO_nc
				opcao="C"
				else
					if nota ="TB_NOTA_D" then
					CAMINHOn = CAMINHO_nd
					opcao="D"
					else
						if nota ="TB_NOTA_E" then
						CAMINHOn = CAMINHO_ne
						opcao="E"
						else
							if nota ="TB_NOTA_F" then
								CAMINHOn = CAMINHO_nf	
								opcao="F"	
							else	
								if nota ="TB_NOTA_V" then
									CAMINHOn = CAMINHO_nv	
									opcao="V"
								else
									if nota ="TB_NOTA_K" then
										CAMINHOn = CAMINHO_nk	
										opcao="K"
									else
										response.Write("ERRO")								
									End if						
								End if									
							End if								
						End if
					End if
				End if
			end if
		end if	
		%>																		
		<select name="avaliacoes" class="select_style" id="avaliacoes" onChange="MM_callJS('submitfuncao()')">
											  <option value="999990"></option>
		<%
		
		dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,outro)
			dados_separados=split(dados_tabela,"#$#")
			ln_nom_cols=dados_separados(4)
			nm_vars=dados_separados(5)
			nm_bd=dados_separados(6)
			avaliacoes_nomes=split(ln_nom_cols,"#!#")
			verifica_avaliacoes=split(nm_vars,"#!#")
			avaliacoes=split(nm_bd,"#!#")
		
		for i=2 to UBOUND(avaliacoes_nomes)
			j=i-2
			if avaliacoes(j)="CALCULADO" or  avaliacoes(j)="rs" or avaliacoes(j)="rb" then
			else
		%>
											  <option value="<%response.Write(avaliacoes(j))%>"> 
											  <%response.Write(avaliacoes_nomes(i))%>
		  </option>
		<%
			end if
		NEXT
		%>
		</select> 
	<%
	end if
elseif opt="av2" then

session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")

%>																		
<select name="avaliacoes" class="select_style" id="avaliacoes" onChange="MM_callJS('submitfuncao()')">
                                      <option value="999990"></option> 
<%

		Set RSTB = Server.CreateObject("ADODB.Recordset")
		CONEXAO = "Select * from TB_Da_Aula WHERE NU_Unidade = "& session("u_pub") &" AND CO_Curso = '"& session("c_pub") &"' AND CO_Etapa = '"& session("e_pub")  &"'"
		Set RSTB = CONg.Execute(CONEXAO)

'response.Write(CONEXAO)

nota = RSTB("TP_Nota")

if nota ="TB_NOTA_A" then
CAMINHOn = CAMINHO_na
opcao="A"
else
	if nota="TB_NOTA_B" then
	CAMINHOn = CAMINHO_nb
	opcao="B"
	else
		if nota ="TB_NOTA_C" then
		CAMINHOn = CAMINHO_nc
		opcao="C"
		else
			if nota ="TB_NOTA_D" then
			CAMINHOn = CAMINHO_nd
			opcao="D"
			else
				if nota ="TB_NOTA_E" then
				CAMINHOn = CAMINHO_ne
				opcao="E"
				else
					if nota ="TB_NOTA_F" then
						CAMINHOn = CAMINHO_nf	
						opcao="F"	
					else	
						if nota ="TB_NOTA_V" then
							CAMINHOn = CAMINHO_nv	
							opcao="V"
						else
							if nota ="TB_NOTA_K" then
								CAMINHOn = CAMINHO_nk	
								opcao="K"
							else
								response.Write("ERRO")								
							End if						
						End if									
					End if	
				End if
			End if
		End if
	end if
end if	

dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,outro)
	dados_separados=split(dados_tabela,"#$#")
	ln_nom_cols=dados_separados(4)
	nm_vars=dados_separados(5)
	nm_bd=dados_separados(6)
	avaliacoes_nomes=split(ln_nom_cols,"#!#")
	verifica_avaliacoes=split(nm_vars,"#!#")
	avaliacoes=split(nm_bd,"#!#")

for i=2 to UBOUND(avaliacoes_nomes)
	j=i-2
	if avaliacoes(j)="CALCULADO" or  avaliacoes(j)="rs" or avaliacoes(j)="rb" then
	else
%>
                                      <option value="<%response.Write(avaliacoes(j))%>"> 
                                      <%response.Write(avaliacoes_nomes(i))%>
                                      </option>
<%
	end if
NEXT
%>
</select> 
<%
elseif opt="mi" then
'etapa_altera=Request.Form("e_pub")
session("u_pub")=session("u_pub")
session("c_pub")=session("c_pub")
session("e_pub")=session("e_pub")
session("t_pub")=session("t_pub")

%>
<select name="mediainformada" class="select_style" id="mediainformada" onChange="MM_callJS('submitfuncao()')">
<option value="0" selected></option>
<%i=0
while i<101 %>
                                      <option value="<%=i%>"> 
                                      <%response.Write(i)%>
                                      </option>
<%i=i+5
WEND%>
</select>	
<%elseif opt="imgtb" then

matric=Request.Form("matr_tb_pub")

vetor_fotos=Session("vetor_fotos")

	Set RSs = Server.CreateObject("ADODB.Recordset")
	SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& matric
	Set RSs = CON_AL.Execute(SQL_s)

	nome_aluno=RSs("NO_Aluno")	

'vetor_arquivos=split(vetor_fotos,",")

for i = 0 to ubound(vetor_fotos)
nome_arquivo =vetor_fotos(i)
nome_jpg=matric&".jpg"
lowercase=lcase(nome_arquivo)
'response.Write(nome_jpg&"-"&lowercase&"-"&mostra_img)
	if nome_jpg=lowercase then
		mostra_img="OK"
	elseif mostra_img<>"OK" then
		mostra_img="NO"
	end if
next	

if mostra_img="OK" then

%>
<a href="#" title="<% response.Write(Server.URLEncode(nome_aluno)) %>" onClick="centraliza(500,536);MM_showHideLayers('fundo','','show','alinha','','show');mclosetime()"><img src="../../../../img/fotos/aluno/<% response.Write(matric) %>.jpg" alt="<% response.Write(Server.URLEncode(nome_aluno)) %>" width="50" height="60" border="0"></a>

<%end if%>




<%elseif opt="img" then
num_cham=Request.Form("num_cham_pub")
co_matric=Request.Form("matric_pub")
periodo_exibe=Request.Form("periodo_pub")

nom_periodo=split(periodo_exibe,"#!#")

	parametros_chamada_jscript="celula"&num_cham
	for b=1 to ubound(nom_periodo) 
		parametros_chamada_jscript=parametros_chamada_jscript&",celula"&num_cham&"p"&b
	next

	Set RSs = Server.CreateObject("ADODB.Recordset")
	SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& co_matric
	Set RSs = CON_AL.Execute(SQL_s)

	nom_aluno=RSs("NO_Aluno")	
	
%>
   <div id="alinha" style="position:absolute; width:500px; z-index: 3; height: 536px; visibility: hidden;"> 
    <table border="0" cellspacing="0" bgcolor="#FFFFFF">
        <tr> 
          <td height="16"> 
            <div align="right"> <span class="voltar1"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide');focar('<%response.Write(num_cham&"c2")%>');mudar_cor_focus(<%response.Write(parametros_chamada_jscript)%>)">fechar</a>&nbsp;<a href="#" onClick="MM_showHideLayers('fundo','','hide','alinha','','hide');focar('<%response.Write(num_cham&"c2")%>');mudar_cor_focus(<%response.Write(parametros_chamada_jscript)%>)"><img src="../../../../img/fecha.gif" width="20" height="16" border="0" align="absbottom"></a></font></span></div></td>
        </tr>
        <tr> 
          <td><div align="center" ><img src="../../../../img/fotos/aluno/<% Response.Write(co_matric) %>.jpg" height="500"></div></td>
        </tr>
        <tr>
          <td height="20">
    <div align="center"><font size="1" face="Verdana, Arial, Helvetica, sans-serif"> 
              <% Response.Write(Server.URLEncode(nom_aluno)) %>
              </font></div></td>
        </tr>
      </table>
    </div>



<%elseif opt="nt" then

co_matric=Request.Form("matric_pub")
CAMINHOn=Request.Form("caminho_pub")
opcao=Request.Form("opcao_pub")
periodo=Request.Form("outro_pub")
ano_letivo=Request.Form("ano_pub")
wrk_materia=Request.Form("materia_pub")
curso=Request.Form("c_pub")
co_etapa=Request.Form("e_pub")
larg_max=Request.Form("larg_max_pub")

if periodo<>0 then
	sigla_periodo=periodos(periodo,"sigla")	
end if
Set CON_N = Server.CreateObject("ADODB.Connection")
ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
CON_N.Open ABRIR3

Set RS = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& wrk_materia &"'"
RS.Open SQL, CON0

mae= RS("IN_MAE")
fil= RS("IN_FIL")
in_co= RS("IN_CO")
peso= RS("NU_Peso")

if (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE) or (mae=TRUE and fil=FALSE and in_co=TRUE) then

else

	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT * FROM TB_Materia where CO_Materia ='"& wrk_materia &"' order by NU_Ordem_Boletim"
	RS1.Open SQL1, CON0
	
	if RS1.EOF then
		response.Write("ERRO na sele&ccedil;&atilde;o de mat&eacute;ria - Funcoes 6 - ln 965")
	else
		wrk_mat_princ=RS1("CO_Materia_Principal")
		if wrk_mat_princ="" or isnull(wrk_mat_princ) or wrk_mat_princ= " " then
			wrk_co_materia=wrk_materia
			wrk_mat_princ=wrk_materia
		else
			wrk_co_materia=wrk_materia
		end if
		
'CAMINHOn = CAMINHOn.replace(/\+/g," ")
'CAMINHOn = unescape(CAMINHOn)
'response.Write(CAMINHOn)

	dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,periodo)
	dados_separados=split(dados_tabela,"#$#")
	tb=dados_separados(0)
	ln_pesos_cols=dados_separados(1)
	ln_pesos_vars=dados_separados(2)
	nm_pesos_vars=dados_separados(3)
	ln_nom_cols=dados_separados(4)
	nm_vars=dados_separados(5)
	nm_bd=dados_separados(6)
	vars_calc=dados_separados(7)
	action=dados_separados(8)
	notas_a_lancar=dados_separados(9)

	linha_pesos=split(ln_pesos_cols,"#!#")
	linha_pesos_variaveis=split(ln_pesos_vars,"#!#")
	nome_pesos_variaveis=split(nm_pesos_vars,"#!#")
	linha_nome_colunas=split(ln_nom_cols,"#!#")
	nome_variaveis=split(nm_vars,"#!#")
	variaveis_bd=split(nm_bd,"#!#")
	calcula_variavel=split(vars_calc,"#!#")
	
qtd_colunas=UBOUND(linha_nome_colunas)+1
width_num_cham="20"
width_nom_aluno="40"
width_else=(larg_max-width_num_cham-width_nom_aluno)/(qtd_colunas-2)	
%>
<table width="<%response.Write(larg_max)%>" border="0" cellspacing="0" cellpadding="0">
  <tr> 

 <% for i= 0 to ubound(linha_pesos)
 		if i=0 then
			width=width_num_cham
			align="center"
		elseif i=1 then	
			width=width_nom_aluno
			align="left"			
		else
			width=width_else
			align="center"
		end if		

		if linha_pesos(i)="PESO" then
			linha_pesos_variaveis=split(ln_pesos_vars,"#!#")
			nome_pesos_variaveis=split(nm_pesos_vars,"#!#")		
						
				Set RSpeso = Server.CreateObject("ADODB.Recordset")
				SQL_peso = "Select "&linha_pesos_variaveis(i)&" from "& tb &" WHERE CO_Matricula = "& co_matric & " AND CO_Materia_Principal = '"& wrk_mat_princ &"' AND CO_Materia = '"& wrk_co_materia &"' AND NU_Periodo="&periodo
				Set RSpeso = CON_N.Execute(SQL_peso)
					
				if RSpeso.EOF then
				else	
					valor_peso=RSpeso(""&linha_pesos_variaveis(i)&"")
				end if		
				IF comunica="s" THEN		
			'		linha_pesos(i)=valor_peso&"<input name="&nome_pesos_variaveis(i)&" type=""hidden"" id="&nome_pesos_variaveis(i)&" class=""peso"" value="&valor_peso&">"	
					linha_pesos(i)=valor_peso
				else	
			'		linha_pesos(i)="<input name="&nome_pesos_variaveis(i)&" type="&tipo&" id="&nome_pesos_variaveis(i)&" class=""peso"" value="&valor_peso&">"
					linha_pesos(i)=valor_peso
				end if	
			
		end if				
 %>
    <td width="<%response.Write(width)%>" class="<%response.Write(classe_subtit)%>"><div align="<%response.Write(align)%>"><%response.Write(linha_pesos(i))%></div></td>
<%	next%>
</tr>
  <tr> 
 <% for j= 0 to ubound(linha_nome_colunas)
 		if j=0 then
			width=width_num_cham
			align="center"
		elseif j=1 then	
			width=width_nom_aluno
			align="left"			
		else
			width=width_else
			align="center"
		end if		
	if linha_nome_colunas(j)="N&ordm;" then
		cabecalho="Per"
	elseif linha_nome_colunas(j)="Nome" then
		cabecalho="Disciplina"
	else
		cabecalho=linha_nome_colunas(j)
	end if	
 %>
    <td width="<%response.Write(width)%>" class="<%response.Write(classe_subtit)%>"><div align="<%response.Write(align)%>"><%response.Write(cabecalho)%></div></td>
<%	next%>  
  </tr>
  <%
	if periodo = 0 then
	qtd_cols_this_if=ubound(nome_variaveis)+3
	%>
        <tr>
            <td width="<%response.Write(width)%>" colspan="<%response.Write(qtd_cols_this_if)%>" class="<%response.Write(classe)%>">N&atilde;o Dispon&iacute;vel para esse per&iacute;odo.</td>                                      						
  </tr>
	<%
    else
		Set RSs = Server.CreateObject("ADODB.Recordset")
		SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& co_matric&" and TB_Matriculas.NU_Ano="&ano_letivo
		Set RSs = CON_AL.Execute(SQL_s)
	
		if RSs.EOF then
		%>
				<tr>
					<td width="<%response.Write(width)%>"><%response.Write(sigla_periodo)%></td>
					<td width="<%response.Write(width)%>"><%response.Write(wrk_co_materia)%></td>                                      						
					<%for m= 0 to ubound(nome_variaveis)
						width=width_else
						align="center"
				 %>
						<td width="<%response.Write(width)%>" class="<%response.Write(classe)%>">&nbsp;</td>
					 <%next%>
  </tr>
		 <%else		
			%>   
				<tr>
					<td width="<%response.Write(width)%>"><%response.Write(sigla_periodo)%></td>
					<td width="<%response.Write(width)%>" ><%response.Write(wrk_co_materia)%></td>               
					 <% 
					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& co_matric & " AND CO_Materia_Principal = '"& wrk_mat_princ &"' AND CO_Materia = '"& wrk_co_materia &"' AND NU_Periodo="&periodo
					Set RS3 = CON_N.Execute(SQL_N)			 
					coluna=0	 
					 for n= 0 to ubound(nome_variaveis)
						width=width_else
						align="center"
						if RS3.EOF then 
							valor=""
						else
							if variaveis_bd(n)="CALCULADO" then
								valor="&nbsp;"
								'Nesse caso o valor é calculado pela função calcular_nota chamada mais abaixo
							else
								valor=RS3(""&variaveis_bd(n)&"")
							end if					
						end if
						
						if (valor="" or isnull(valor)) then
							coluna=coluna+1	
							conteudo="&nbsp;"			
						else
							if nome_variaveis(n)="faltas" then
								coluna=coluna+1
								conteudo=formatnumber(valor,0)		
							elseif nome_variaveis(n)="nav" or nome_variaveis(n)="media_teste" or nome_variaveis(n)="media_prova" or nome_variaveis(n)="media1" or nome_variaveis(n)="media2" or nome_variaveis(n)="media3" or situac<>"C" then
								coluna=coluna	
								conteudo=formatnumber(valor,0)
							elseif calcula_variavel(n)="CALC1" and situac="C" then
								coluna=coluna
								valor=calcular_nota(calcula_variavel(n),CAMINHOn,tb,nu_matricula,mat_princ,co_materia,periodo)
								conteudo=formatnumber(valor,0)	
							else
								coluna=coluna+1
								conteudo=formatnumber(valor,0)			
							end if	
						end if	
						'conteudo=n
				 %>
					<td width="<%response.Write(width)%>">
						<div align="<%response.Write(align)%>">
							<%response.Write(conteudo)%> 
						</div>
				  </td>
				  <%	next  
				  %>
  </tr>
				  <%			
			end if	
		end if	
	END IF	
END IF	
%>     
</td>
</tr>
</table>
<%elseif opt="ntzoom" then

co_matric=Request.Form("matric_pub")
CAMINHOn=Request.Form("caminho_pub2")
opcao=Request.Form("opcao_pub")
periodo=Request.Form("outro_pub")
ano_letivo=Request.Form("ano_pub")
wrk_materia=Request.Form("materia_pub")
curso=Request.Form("c_pub")
co_etapa=Request.Form("e_pub")
larg_max=Request.Form("larg_max_pub")
CAMINHOn=replace(CAMINHOn,"$b$","\")
CAMINHOn=replace(CAMINHOn,"$u$","_")

if periodo<>0 then
	sigla_periodo=periodos(periodo,"sigla")	
end if
Set CON_N = Server.CreateObject("ADODB.Connection")
ABRIR3 = "DBQ="& CAMINHOn & ";Driver={Microsoft Access Driver (*.mdb)}"
CON_N.Open ABRIR3

Set RS = Server.CreateObject("ADODB.Recordset")
SQL = "SELECT * FROM TB_Programa_Aula where CO_Etapa ='"& co_etapa &"' AND CO_Curso ='"& curso &"' AND CO_Materia ='"& wrk_materia &"'"
RS.Open SQL, CON0

mae= RS("IN_MAE")
fil= RS("IN_FIL")
in_co= RS("IN_CO")
peso= RS("NU_Peso")



if (mae=TRUE and fil=TRUE and in_co=FALSE and isnull(peso)) or (mae=TRUE and fil=TRUE and in_co=FALSE) or (mae=TRUE and fil=FALSE and in_co=TRUE) then

else

	Set RS1 = Server.CreateObject("ADODB.Recordset")
	SQL1 = "SELECT * FROM TB_Materia where CO_Materia ='"& wrk_materia &"' order by NU_Ordem_Boletim"
	RS1.Open SQL1, CON0
	
	if RS1.EOF then
		response.Write("ERRO na sele&ccedil;&atilde;o de mat&eacute;ria - Funcoes 6 - ln 965")
	else
		wrk_mat_princ=RS1("CO_Materia_Principal")
		if wrk_mat_princ="" or isnull(wrk_mat_princ) or wrk_mat_princ= " " then
			wrk_co_materia=wrk_materia
			wrk_mat_princ=wrk_materia
		else
			wrk_co_materia=wrk_materia
		end if
		

	dados_tabela=verifica_dados_tabela(CAMINHOn,opcao,periodo)
	dados_separados=split(dados_tabela,"#$#")
	tb=dados_separados(0)
	ln_pesos_cols=dados_separados(1)
	ln_pesos_vars=dados_separados(2)
	nm_pesos_vars=dados_separados(3)
	ln_nom_cols=dados_separados(4)
	nm_vars=dados_separados(5)
	nm_bd=dados_separados(6)
	vars_calc=dados_separados(7)
	action=dados_separados(8)
	notas_a_lancar=dados_separados(9)

	linha_pesos=split(ln_pesos_cols,"#!#")
	linha_pesos_variaveis=split(ln_pesos_vars,"#!#")
	nome_pesos_variaveis=split(nm_pesos_vars,"#!#")
	linha_nome_colunas=split(ln_nom_cols,"#!#")
	nome_variaveis=split(nm_vars,"#!#")
	variaveis_bd=split(nm_bd,"#!#")
	calcula_variavel=split(vars_calc,"#!#")
	
qtd_colunas=UBOUND(linha_nome_colunas)+1
width_num_cham="20"
width_nom_aluno="40"
width_else=(larg_max-width_num_cham-width_nom_aluno)/(qtd_colunas-2)	
%>
<table width="<%response.Write(larg_max)%>" border="0" cellspacing="0" cellpadding="0">
  <tr> 

 <% for i= 0 to ubound(linha_pesos)
 		if i=0 then
			width=width_num_cham
			align="center"
		elseif i=1 then	
			width=width_nom_aluno
			align="left"			
		else
			width=width_else
			align="center"
		end if		

		if linha_pesos(i)="PESO" then
			linha_pesos_variaveis=split(ln_pesos_vars,"#!#")
			nome_pesos_variaveis=split(nm_pesos_vars,"#!#")		
						
				Set RSpeso = Server.CreateObject("ADODB.Recordset")
				SQL_peso = "Select "&linha_pesos_variaveis(i)&" from "& tb &" WHERE CO_Matricula = "& co_matric & " AND CO_Materia_Principal = '"& wrk_mat_princ &"' AND CO_Materia = '"& wrk_co_materia &"' AND NU_Periodo="&periodo
				Set RSpeso = CON_N.Execute(SQL_peso)
					
				if RSpeso.EOF then
				else	
					valor_peso=RSpeso(""&linha_pesos_variaveis(i)&"")
				end if		
				IF comunica="s" THEN		
'					linha_pesos(i)=valor_peso&"<input name="&nome_pesos_variaveis(i)&" type=""hidden"" id="&nome_pesos_variaveis(i)&" class=""peso"" value="&valor_peso&">"	
					linha_pesos(i)=valor_peso
				else	
				'	linha_pesos(i)="<input name="&nome_pesos_variaveis(i)&" type="&tipo&" id="&nome_pesos_variaveis(i)&" class=""peso"" value="&valor_peso&">"
					linha_pesos(i)=valor_peso
				end if	
			
		end if				
 %>
    <td width="<%response.Write(width)%>" class="zoom_label"><div align="<%response.Write(align)%>"><%response.Write(linha_pesos(i))%></div></td>
<%	next%>
</tr>
  <tr> 
 <% for j= 0 to ubound(linha_nome_colunas)
 		if j=0 then
			width=width_num_cham
			align="center"
		elseif j=1 then	
			width=width_nom_aluno
			align="left"			
		else
			width=width_else
			align="center"
		end if		
	if linha_nome_colunas(j)="N&ordm;" then
		cabecalho="Per"
	elseif linha_nome_colunas(j)="Nome" then
		cabecalho="Disciplina"
	else
		cabecalho=linha_nome_colunas(j)
	end if	
 %>
    <td width="<%response.Write(width)%>" class="zoom_tit"><div align="<%response.Write(align)%>"><%response.Write(cabecalho)%></div></td>
<%	next%>  
  </tr>
  <%
	if periodo = 0 then
	qtd_cols_this_if=ubound(nome_variaveis)+3
	%>
        <tr>
            <td width="<%response.Write(width)%>" colspan="<%response.Write(qtd_cols_this_if)%>" class="zoom_texto">N&atilde;o Dispon&iacute;vel para esse per&iacute;odo.</td>                                      						
  </tr>
	<%
    else
		Set RSs = Server.CreateObject("ADODB.Recordset")
		SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& co_matric&" and TB_Matriculas.NU_Ano="&ano_letivo
		Set RSs = CON_AL.Execute(SQL_s)
	
		if RSs.EOF then
		%>
				<tr>
					<td width="<%response.Write(width)%>" class="zoom_texto"><%response.Write(sigla_periodo)%></td>
					<td width="<%response.Write(width)%>" class="zoom_texto"><%response.Write(wrk_co_materia)%></td>                                      						
					<%for m= 0 to ubound(nome_variaveis)
						width=width_else
						align="center"
				 %>
						<td width="<%response.Write(width)%>" class="zoom_texto">&nbsp;</td>
					 <%next%>
  </tr>
		 <%else		
			%>   
				<tr>
					<td width="<%response.Write(width)%>" class="zoom_texto"><%response.Write(sigla_periodo)%></td>
					<td width="<%response.Write(width)%>" class="zoom_texto"><%response.Write(wrk_co_materia)%></td>               
					 <% 
					Set RS3 = Server.CreateObject("ADODB.Recordset")
					SQL_N = "Select * from "& tb &" WHERE CO_Matricula = "& co_matric & " AND CO_Materia_Principal = '"& wrk_mat_princ &"' AND CO_Materia = '"& wrk_co_materia &"' AND NU_Periodo="&periodo
					Set RS3 = CON_N.Execute(SQL_N)			 
					coluna=0	 
					 for n= 0 to ubound(nome_variaveis)
						width=width_else
						align="center"
						if RS3.EOF then 
							valor=""
						else
							if variaveis_bd(n)="CALCULADO" then
								valor="&nbsp;"
								'Nesse caso o valor é calculado pela função calcular_nota chamada mais abaixo
							else
								valor=RS3(""&variaveis_bd(n)&"")
							end if					
						end if
						
						if (valor="" or isnull(valor)) then
							coluna=coluna+1	
							conteudo="&nbsp;"			
						else
							if nome_variaveis(n)="faltas" then
								coluna=coluna+1
								conteudo=formatnumber(valor,0)		
							elseif nome_variaveis(n)="nav" or nome_variaveis(n)="media_teste" or nome_variaveis(n)="media_prova" or nome_variaveis(n)="media1" or nome_variaveis(n)="media2" or nome_variaveis(n)="media3" or situac<>"C" then
								coluna=coluna	
								conteudo=formatnumber(valor,0)
							elseif calcula_variavel(n)="CALC1" and situac="C" then
								coluna=coluna
								valor=calcular_nota(calcula_variavel(n),CAMINHOn,tb,nu_matricula,mat_princ,co_materia,periodo)
								conteudo=formatnumber(valor,0)	
							else
								coluna=coluna+1
								conteudo=formatnumber(valor,0)			
							end if	
						end if	
						'conteudo=n
				 %>
					<td width="<%response.Write(width)%>" class="zoom_texto">
						<div align="<%response.Write(align)%>">
							<%response.Write(conteudo)%> 
						</div>
				  </td>
				  <%	next  
				  %>
  </tr>
				  <%			
			end if	
		end if	
	END IF	
END IF	
%>     
</td>
</tr>
</table>
<%elseif opt="ocr" then
co_matric=Request.Form("matric_pub")
		
	Set RS4 = Server.CreateObject("ADODB.Recordset")
	SQL4 = "select * from TB_Ocorrencia_Aluno where CO_Matricula = " & co_matric &" Order BY DA_Ocorrencia DESC,HO_Ocorrencia"
	set RS4 = CON3.Execute (SQL4)
	if RS4.EOF then
	%>
	<div align="left"><%Response.Write(Server.URLEncode("&nbsp;&nbsp;Última Ocorrencia Registrada:&nbsp;Sem Ocorrências"))%></div>
	<%else
	data_ocor=RS4("DA_Ocorrencia")
	ho_ocor=RS4("HO_Ocorrencia")
	assunto=RS4("CO_Assunto")
	co_ocr=RS4("CO_Ocorrencia")
	prof=RS4("CO_Professor")
	mat_ocr=RS4("NO_Materia")
	obs=RS4("TX_Observa")	
	
	Set RSto = Server.CreateObject("ADODB.Recordset")
	SQLto = "SELECT * FROM TB_Tipo_Ocorrencia Where CO_Ocorrencia="&co_ocr
	RSto.Open SQLto, CON0	
	no_ocorrencia=RSto("NO_Ocorrencia")
	%>
	<div align="left"><%Response.Write(Server.URLEncode("&nbsp;&nbsp;Última Ocorrencia Registrada:&nbsp;"&data_ocor&", referente a: "&no_ocorrencia))%></div>
	<%end if

elseif opt="nm" then
co_matric=Request.Form("matric_pub")
		
	Set RSs = Server.CreateObject("ADODB.Recordset")
	SQL_s ="SELECT * FROM TB_Alunos INNER JOIN TB_Matriculas ON TB_Alunos.CO_Matricula=TB_Matriculas.CO_Matricula where TB_Matriculas.CO_Matricula ="& co_matric&" and TB_Matriculas.NU_Ano="&ano_letivo
	Set RSs = CON_AL.Execute(SQL_s)


	nome_aluno=RSs("NO_Aluno")
	
	if RSs.EOF then
	%>
	<div align="left"><%Response.Write(Server.URLEncode("executa.asp>1200-Erro na busca do nome"))%></div>
	<%else
		Set RSCONTA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT * FROM TB_Contatos WHERE TP_Contato ='ALUNO' And CO_Matricula ="& co_matric
		RSCONTA.Open SQLA, CONCONT

		if RSCONTA.EOF then
			nascimento="0/0/0"
		else
			nascimento = RSCONTA("DA_Nascimento_Contato")
		end if
			vetor_nascimento = Split(nascimento,"/")  
			dia = vetor_nascimento(0)
			mes = vetor_nascimento(1)
			ano = vetor_nascimento(2)
			
			data= dia&"-"&mes&"-"&ano
			intervalo = DateDiff("d", data , now )
			
			idade = int(intervalo/365.25)

	%>
	<div align="left"><%Response.Write("&nbsp;&nbsp;"&Server.URLEncode(nome_aluno)&"&nbsp;-&nbsp;"&idade&" anos")%></div>
	<%end if
elseif opt="ctrl" then
session("u_pub")=session("u_pub")
session("c_pub")=Request.Form("c_pub")

			

		Set RS0b = Server.CreateObject("ADODB.Recordset")
		SQL0b = "SELECT DISTINCT CO_Etapa FROM TB_Unidade_Possui_Etapas where NU_Unidade="&session("u_pub")&"AND CO_Curso='"&session("c_pub")&"'"
		RS0b.Open SQL0b, CON0
				


%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
<%
While not RS0b.EOF
co_etapa = RS0b("CO_Etapa")

		Set RS0c = Server.CreateObject("ADODB.Recordset")
		SQL0c = "SELECT * FROM TB_Etapa where CO_Curso='"&session("c_pub")&"' AND CO_Etapa='"&co_etapa&"'"
		RS0c.Open SQL0c, CON0
		
	no_etapa = RS0c("NO_Etapa")		
	
	
		Set RS_WF = Server.CreateObject("ADODB.Recordset")
		SQL_WF = "SELECT * FROM TB_Autoriza_WF WHERE NU_Unidade="&session("u_pub")&" AND CO_Curso='"&session("c_pub")&"' and CO_Etapa='"&co_etapa&"'"
		RS_WF.Open SQL_WF, CON_WF
					
		co_apr1=RS_WF("CO_apr1")
		co_apr2=RS_WF("CO_apr2")
		co_apr3=RS_WF("CO_apr3")
		co_apr4=RS_WF("CO_apr4")
		co_apr5=RS_WF("CO_apr5")
		co_apr6=RS_WF("CO_apr6")
		co_apr7=RS_WF("CO_apr7")					
		co_prova1=RS_WF("CO_prova1")
		co_prova2=RS_WF("CO_prova2")
		co_prova3=RS_WF("CO_prova3")
		co_prova4=RS_WF("CO_prova4")
		co_prova5=RS_WF("CO_prova5")		
		co_prova6=RS_WF("CO_prova6")	
		co_prova7=RS_WF("CO_prova7")
	
	
	
%>
<tr><td class="tb_tit"><%response.Write(Server.URLEncode(no_etapa))%></td></tr>	
  <tr>
    <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
                    <tr>
                    <% 
					
					Set RSP = Server.CreateObject("ADODB.Recordset")
					SQLP = "SELECT Distinct(NU_Periodo),NO_Periodo FROM TB_Periodo order by NU_Periodo"
					RSP.Open SQLP, CON0
							
					conta_periodo=0
					while not RSP.EOF				
						nome_periodo=RSP("NO_Periodo")
					%>
                      <td colspan="2" class="tb_subtit"><div align="center"><%response.Write(Server.URLEncode(nome_periodo))%></div></td>
					<%
					conta_periodo=conta_periodo+1
					RSP.MOVENEXT
					WEND
					%>
						
                      </tr>
                    <tr>
                    <% for cp=1 to conta_periodo %>
                      <td width="71" class="tb_subtit"><div align="center">Testes 
                      </div></td>
                      <td width="72" class="tb_subtit"><div align="center">Provas 
                      </div></td>
					<%next%>
                      </tr>
                    <tr>
                    <% 

											
					for cp=1 to conta_periodo 	
										
						if cp=1 then						
							if co_apr1="L" then
								apr_checked_lib ="checked"
								apr_checked_blq =""														
							else
								apr_checked_lib =""
								apr_checked_blq ="checked"																	
							end if
							if co_prova1="L" then
								pr_checked_lib ="checked"
								pr_checked_blq =""														
							else
								pr_checked_lib =""
								pr_checked_blq ="checked"																	
							end if			
											
						elseif cp=2 then
							if co_apr2="L" then
								apr_checked_lib	="checked"
								apr_checked_blq	=""														
							else
								apr_checked_lib	=""
								apr_checked_blq	="checked"																	
							end if
							if co_prova2="L" then
								pr_checked_lib	="checked"
								pr_checked_blq	=""														
							else
								pr_checked_lib	=""
								pr_checked_blq	="checked"																	
							end if												
						elseif cp=3 then
							if co_apr3="L" then
								apr_checked_lib	="checked"
								apr_checked_blq	=""														
							else
								apr_checked_lib	=""
								apr_checked_blq	="checked"																	
							end if
							if co_prova3="L" then
								pr_checked_lib	="checked"
								pr_checked_blq	=""														
							else
								pr_checked_lib	=""
								pr_checked_blq	="checked"																	
							end if	
					
						elseif cp=4 then
							if co_apr4="L" then
								apr_checked_lib	="checked"
								apr_checked_blq	=""														
							else
								apr_checked_lib	=""
								apr_checked_blq	="checked"																	
							end if
							if co_prova4="L" then
								pr_checked_lib	="checked"
								pr_checked_blq	=""														
							else
								pr_checked_lib	=""
								pr_checked_blq	="checked"																	
							end if					
				
						elseif cp=5 then
							if co_apr5="L" then
								apr_checked_lib	="checked"
								apr_checked_blq	=""														
							else
								apr_checked_lib	=""
								apr_checked_blq	="checked"																	
							end if
							if co_prova5="L" then
								pr_checked_lib	="checked"
								pr_checked_blq	=""														
							else
								pr_checked_lib	=""
								pr_checked_blq	="checked"																	
							end if						
						elseif cp=6 then
							if co_apr6="L" then
								apr_checked_lib	="checked"
								apr_checked_blq	=""														
							else
								apr_checked_lib	=""
								apr_checked_blq	="checked"																	
							end if
							if co_prova6="L" then
								pr_checked_lib	="checked"
								pr_checked_blq	=""														
							else
								pr_checked_lib	=""
								pr_checked_blq	="checked"																	
							end if		
						elseif cp=7 then
							if co_apr7="L" then
								apr_checked_lib	="checked"
								apr_checked_blq	=""														
							else
								apr_checked_lib	=""
								apr_checked_blq	="checked"																	
							end if
							if co_prova7="L" then
								pr_checked_lib	="checked"
								pr_checked_blq	=""														
							else
								pr_checked_lib	=""
								pr_checked_blq	="checked"																	
							end if	
						end if												
					%>                    
                      <td width="71"><div align="center">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td>
                            <input type="radio" name="<%response.Write(co_etapa&"_t_"&cp)%>" id="<%response.Write(co_etapa&"_t_"&cp)%>" value="D" class="borda" <%response.Write(apr_checked_blq)%> onclick="GravaControle('<%response.Write(co_etapa&"_t_"&cp)%>',this.value)"/></td>
                            <td class="form_dado_texto">Bloq</td>
                          </tr>
                          <tr>
                            <td><input type="radio" name="<%response.Write(co_etapa&"_t_"&cp)%>" id="<%response.Write(co_etapa&"_t_"&cp)%>" value="L" class="borda" <%response.Write(apr_checked_lib)%> onclick="GravaControle('<%response.Write(co_etapa&"_t_"&cp)%>',this.value)"/></td>
                            <td class="form_dado_texto">Lib</td>
                          </tr>
                        </table>
                      </div></td>
                      <td width="72"><div align="center">
                        <table width="100%" border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td><input type="radio" name="<%response.Write(co_etapa&"_p_"&cp)%>" id="<%response.Write(co_etapa&"_p_"&cp)%>" value="D" class="borda" <%response.Write(pr_checked_blq)%> onclick="GravaControle('<%response.Write(co_etapa&"_p_"&cp)%>',this.value)"/></td>
                            <td class="form_dado_texto">Bloq</td>
                          </tr>
                          <tr>
                            <td><input type="radio" name="<%response.Write(co_etapa&"_p_"&cp)%>" id="<%response.Write(co_etapa&"_p_"&cp)%>" value="L" class="borda" <%response.Write(pr_checked_lib)%> onclick="GravaControle('<%response.Write(co_etapa&"_p_"&cp)%>',this.value)"/></td>
                            <td class="form_dado_texto">Lib</td>
                          </tr>
                        </table>
                      </div></td>                   
                      <%next%>

                    </tr>
                    
                  </table>               
                  </td>
  </tr>
<%RS0b.MOVENEXT
WEND
%>    
</table>


<%else

session("t_pub")=Request.Form("t_pub")
end if%>
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
'response.redirect("../../../../inc/erro.asp")
end if
%>