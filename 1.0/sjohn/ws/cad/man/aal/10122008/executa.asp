<%'On Error Resume Next%>
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->
<!--#include file="../../../../inc/caminhos.asp"-->

<% 
opt= request.querystring("opt")
vinculado= request.form("vinc_pub")
aluno= request.form("aluno_pub")


		Set CON = Server.CreateObject("ADODB.Connection") 
		ABRIR = "DBQ="& CAMINHO & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON.Open ABRIR		

		Set CON1 = Server.CreateObject("ADODB.Connection") 
		ABRIR1 = "DBQ="& CAMINHO_al & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1.Open ABRIR1

		Set CON1_aux = Server.CreateObject("ADODB.Connection") 
		ABRIR1_aux = "DBQ="& CAMINHO_al_aux & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON1_aux.Open ABRIR1_aux		
		
		Set CONCONT = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT = "DBQ="& CAMINHO_ct & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT.Open ABRIRCONT
		
		Set CONCONT_aux = Server.CreateObject("ADODB.Connection") 
		ABRIRCONT_aux = "DBQ="& CAMINHO_ct_aux & ";Driver={Microsoft Access Driver (*.mdb)}"
		CONCONT_aux.Open ABRIRCONT_aux		
		
		Set CON0 = Server.CreateObject("ADODB.Connection") 
		ABRIR0 = "DBQ="& CAMINHO_pr & ";Driver={Microsoft Access Driver (*.mdb)}"
		CON0.Open ABRIR0


		Set RSCONTPR = Server.CreateObject("ADODB.Recordset")
		SQLCONTPR = "SELECT * FROM TB_Tipo_Contatos Order by NU_Prioridade_Combo"
		RSCONTPR.Open SQLCONTPR, CON0
		
		
total_tp_familiares=0		
while not RSCONTPR.EOF
tp_resp = RSCONTPR("TP_Contato")
no_tp_resp = RSCONTPR("TX_Descricao")
ordem_familiares=ordem_familiares&"##"&tp_resp&"!!"&no_tp_resp
total_tp_familiares=total_tp_familiares+1

if total_tp_familiares=1 then
foco_default=tp_resp
end if

RSCONTPR.MOVENEXT
WEND

if opt="d" then
		'response.Write("DELETE * FROM TBI_Contatos WHERE CO_Matricula ="&aluno&" and CO_Matricula_Vinc="&vinculado)

		Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
		SQLAC_delete= "DELETE * FROM TBI_Contatos WHERE CO_Matricula ="&aluno&" and CO_Matricula_Vinc="&vinculado
		RSCONTATO_aux_delete.Open SQLAC_delete, CONCONT_aux
		
		Set RSCONTPR = Server.CreateObject("ADODB.Recordset")
		SQLCONTPR = "SELECT * FROM TB_Tipo_Contatos Order by NU_Prioridade_Combo"
		RSCONTPR.Open SQLCONTPR, CON0
foco="nulo"		
familiares=0
pai_cadastrado="n"
mae_cadastrado="n"
while not RSCONTPR.EOF
tp_resp = RSCONTPR("TP_Contato")
no_tp_resp = RSCONTPR("TX_Descricao")
		
		Set RSCONTACONT = Server.CreateObject("ADODB.Recordset")
		SQLAC = "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_resp&"' and CO_Matricula ="& aluno
		RSCONTACONT.Open SQLAC, CONCONT_aux

'response.Write("<br>"&SQLAC)

if RSCONTACONT.EOF then
'response.Write(">>"&familiares)
foco=foco
else

	tipo_contato=RSCONTACONT("TP_Contato")
		if familiares=0 then
		foco=tipo_contato
		end if
		
		if tipo_contato="PAI" OR tipo_contato="MAE" OR tipo_contato="ALUNO" then
				if tipo_contato="PAI" then
				pai_cadastrado="s"
				end if				

				if tipo_contato="MAE" then
				mae_cadastrado="s"
				end if
		else
		nome_contato=RSCONTACONT("NO_Contato")
		end if
	familiares=familiares+1		
end if
RSCONTPR.MOVENEXT
WEND
		
%>

<table width="1000" border="0" align="right" cellspacing="0" class="tb_corpo">
    <td class="tb_tit"
>Dados vinculados ao Aluno</td>
          </tr>
          <tr> 
            <td> 
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr bgcolor="#FFFFFF" background="../../../../img/fundo_interno.gif"> 
                          <td width="5%"  height="10"> <div align="left"><font class="form_dado_texto"> 
                              Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                              </strong></font></div></td>
                          <td width="10%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                            </font><font size="2" face="Arial, Helvetica, sans-serif"> 
                            <input name="vincular" type="text" class="borda" id="vincular" size="12">
                            </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                            </font></td>
                          <td width="37%" height="10"> <div align="right"><font class="form_dado_texto"> 
                              </font></div></td>
                          <td width="2%" height="10" ><font size="2" face="Arial, Helvetica, sans-serif">&nbsp; 
                            </font></td>
                          <td width="46%" height="10"><font size="2" face="Arial, Helvetica, sans-serif">  <!--  'O valor de codigo.value vem do arquivo altera.asp-->
                            <input name="Button" type="button" class="borda_bot" id="Submit" value="Vincular" onClick="VincularAluno(vincular.value,codigo.value)">
                            </font> </td>
                        </tr>
                      </table>		  
			  </td>
          </tr>			
  <tr> 
    <td class="tb_tit"
>Endere&ccedil;o Residencial</td>
  </tr>
  <tr> 
    <td height="10"><table width="100%" border="0" cellspacing="0">
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                          <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                          <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                            
            <input name="rua_res" type="text" class="borda" id="rua_res" value="" size="30" maxlength="60">
                            </font></td>
                          <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                          <td width="19" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="196" class="tb_corpo"
><font class="form_corpo"> <font class="form_corpo"> <font class="form_corpo"> 
                            <input name="num_res" type="text" class="borda" id="num_res" value="" size="12" maxlength="10" onBlur="ValidaNumResFam(this.value)">
                            </font></font></font></td>
                          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                          <td width="11" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="149" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
                              
              <input name="comp_res" type="text" class="borda" id="comp_res" value="" size="20" maxlength="30">
                              </font></div></td>
                        </tr>
                        <tr> 
                          <td width="145" height="21" class="tb_corpo"
><font class="form_dado_texto">Estado</font></td>
                          <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                          <td width="217" height="21" class="tb_corpo"
><font class="form_corpo"> <font class="form_corpo"> 
                            <select name="estadores" class="borda" id="select4" onChange="recuperarCidRes(this.value)">
                              <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")

if isnull(uf_res) or uf_res="" then
uf_res="RJ"
end if

if SG_UF = uf_res then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <%response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <% else %>
                                <option value="<%=SG_UF%>"> 
                                <% response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                              <%
end if	
RS2.MOVENEXT
WEND
%>
                            </select>
                            </font> </font></td>
                          <td width="140" height="21" class="tb_corpo"
><font class="form_dado_texto">Cidade</font></td>
                          <td width="19" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                            <div id="cid_res"> 
                              <select name="cidres" class="borda" id="select10" onChange="recuperarBairroRes(estadores.value,this.value)">
                                <%
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='"&uf_res&"' order by NO_Municipio"
		RS2m.Open SQL2m, CON0
		
while not RS2m.EOF						
SG_UF= RS2m("CO_Municipio")
NO_UF= RS2m("NO_Municipio")
if isnull(cid_res) or cid_res="" then
cid_res=6001
end if
if SG_UF = cid_res then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <%response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <% else %>
                                <option value="<%=SG_UF%>"> 
                                <% response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <%
end if	
RS2m.MOVENEXT
WEND
%>
                              </select>
                            </div>
                            </font></td>
                          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Bairro</font></td>
                          <td width="11" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="149" height="21" class="tb_corpo"
> <div id="bairro_res"><font class="form_corpo"> 
                              <select name="bairrores" class="borda" id="select3">
                                <%
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Municipio="&cid_res&" AND SG_UF='"&uf_res&"' order by NO_Bairro"
		RS2b.Open SQL2b, CON0
		
while not RS2b.EOF						
SG_UF= RS2b("CO_Bairro")
NO_UF= RS2b("NO_Bairro")
if SG_UF=bairro_res then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <%response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <% else %>
                                <option value="<%=SG_UF%>"> 
                                <% response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <%
end if
RS2b.MOVENEXT
WEND
%>
                              </select>
                              </font></div></td>
                        </tr>
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
><font class="form_dado_texto">CEP</font></td>
                          <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                          <td width="217" height="10" class="tb_corpo"
><font class="form_dado_texto"> 
                            <input name="cep" type="text" class="borda" id="cep" onKeyup="formatar(this, '#####-###')" value="" size="11" maxlength="9" onBlur="ValidaCepResFam(this.value)">
                            </font></td>
                          <td width="140" height="10" class="tb_corpo"
>&nbsp;</td>
                          <td width="19" class="tb_corpo"
>&nbsp;</td>
                          <td width="196" class="tb_corpo"
>&nbsp;</td>
                          <td width="90" class="tb_corpo"
>&nbsp;</td>
                          <td width="11" class="tb_corpo"
> <div align="center"></div></td>
                          <td width="149" height="10" class="tb_corpo"
>&nbsp;</td>
                        </tr>
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">Telefones deste endere&ccedil;o</font></div></td>
                          <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                          <td height="10" colspan="2" class="tb_corpo"
><font class="form_corpo"> 
                            
            <input name="tel_res" type="text" class="borda" id="tel_res" value="" size="42" maxlength="100">
                            </font> <div align="left"></div></td>
                          <td width="19" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="196" class="tb_corpo"
>&nbsp;</td>
                          <td width="90" class="tb_corpo"
>&nbsp;</td>
                          <td width="11" class="tb_corpo"
> <div align="center"></div></td>
                          <td width="149" height="10" class="tb_corpo"
>&nbsp;</td>
                        </tr>
                        <tr> 
                          <td height="10" colspan="9" class="tb_tit"
><div align="left">Endere&ccedil;o Comercial </div></td>
                        </tr>
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                          <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                          <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                            
            <input name="rua_com" type="text" class="borda" id="rua_com" value="" size="30" maxlength="60">
                            </font></td>
                          <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                          <td width="19" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                            <input name="num_com" type="text" class="borda" id="num_com" value="" size="12" maxlength="10" onBlur="ValidaNumComFam(this.value)">
                            </font></td>
                          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                          <td class="tb_corpo"
><div align="center">:</div></td>
                          <td width="149" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
                              
              <input name="comp_com" type="text" class="borda" id="comp_com"  value="" size="20" maxlength="30">
                              </font></div></td>
                        </tr>
                        <tr class="tb_corpo"
> 
                          <td width="145" height="26"><font class="form_dado_texto">Estado</font></td>
                          <td width="13"> <div align="left">:</div></td>
                          <td width="217" height="26"><font class="form_corpo"> 
                            <font class="form_corpo"> 
                            <select name="estadocom" class="borda" id="select2" onChange="recuperarCidCom(this.value)">
                              <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")
if isnull(uf_com) or uf_com=""  then
uf_com="RJ"
end if
if SG_UF = uf_com then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <%response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <% else %>
                                <option value="<%=SG_UF%>"> 
                                <% response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                              <%end if						
RS2.MOVENEXT
WEND
%>
                            </select>
                            </font> </font></td>
                          <td width="140" height="26"><font class="form_dado_texto">Cidade</font></td>
                          <td width="19"> <div align="center">:</div></td>
                          <td width="196"> <div id="cid_com"> 
                              <select name="cidcom" class="borda" id="select10" onChange="recuperarBairroCom(estadocom.value,this.value)">
                                <%
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='"&uf_com&"' order by NO_Municipio"
		RS2m.Open SQL2m, CON0
		
while not RS2m.EOF						
SG_UF= RS2m("CO_Municipio")
NO_UF= RS2m("NO_Municipio")
if isnull(cid_com) or cid_com="" then
cid_com=6001
end if
if SG_UF = cid_com then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <%response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <% else %>
                                <option value="<%=SG_UF%>"> 
                                <% response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <%
end if	
RS2m.MOVENEXT
WEND
%>
                              </select>
                            </div></td>
                          <td width="90"><font class="form_dado_texto">Bairro</font></td>
                          <td><div align="center">:</div></td>
                          <td width="149" height="26"> <div id="bairro_com"><font class="form_corpo"> 
                              <select name="bairrocom" class="borda" id="bairro">
                                <%
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Municipio="&cid_com&" AND SG_UF='"&uf_com&"' order by NO_Bairro"
		RS2b.Open SQL2b, CON0
		
while not RS2b.EOF						
SG_UF= RS2b("CO_Bairro")
NO_UF= RS2b("NO_Bairro")

if SG_UF=bairro_com then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <%response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <% else %>
                                <option value="<%=SG_UF%>"> 
                                <% response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <%
end if

RS2b.MOVENEXT
WEND
%>
                              </select>
                              </font></div></td>
                        </tr>
                        <tr class="tb_corpo"
> 
                          <td width="145" height="26"><font class="form_dado_texto">CEP</font></td>
                          <td width="13"> <div align="left">:</div></td>
                          <td width="217" height="26"><font class="form_dado_texto"> 
                            <input name="cep_com" type="text" class="borda" id="cepcom" onKeyup="formatar(this, '#####-###')" value="" size="11" maxlength="9" onBlur="ValidaCepComFam(this.value)">
                            </font></td>
                          <td width="140" height="26">&nbsp;</td>
                          <td width="19">&nbsp;</td>
                          <td width="196">&nbsp;</td>
                          <td width="90">&nbsp;</td>
                          <td><div align="center"></div></td>
                          <td width="149" height="26">&nbsp;</td>
                        </tr>
                        <tr class="tb_corpo"
> 
                          <td width="145" height="28"> <div align="left"><font class="form_dado_texto">Telefones 
                              deste endere&ccedil;o:</font></div></td>
                          <td width="13"> <div align="left">:</div></td>
                          <td height="28" colspan="2"><font class="form_corpo"> 
                            
            <input name="tel_com" type="text" class="borda" id="tel_com" value="" size="42" maxlength="100">
                            </font> <div align="left"></div></td>
                          <td width="19"> <div align="center"></div></td>
                          <td width="196">		  <marquee id="mqLooper1" loop="1"  onStart="<%response.Write("recuperarFamiliares('"&Server.URLEncode(ordem_familiares)&"','"&total_tp_familiares&"','"&foco&"','"&aluno&"','"&aluno&"')")%>"></marquee>
</td>
                          <td width="90">&nbsp;</td>
                          <td><div align="center"></div></td>
                          <td width="149" height="28">&nbsp;</td>
                        </tr>
                      </table></td>
  </tr>
  <tr> 
    <td class="tb_tit"
>Filia&ccedil;&atilde;o</td>
  </tr>
  <tr> 
    <td><table width="100%" border="0" cellspacing="0">
        <tr> 
          <td width="14%" height="26"> <div align="left"><font class="form_dado_texto"> 
              Pai</font></div></td>
          <td width="2%"><div align="center">:</div></td>
          <td width="22%" height="26"><font class="form_corpo"> 
                            <input name="pai" type="text" class="borda" onBlur="recuperarPai(this.value,'p','<%response.Write(pai_cadastrado)%>','<%response.Write(aluno)%>')" value="<%response.Write(pai)%>" size="30" maxlength="50">

            </font></td>
          <td width="15%" height="26"> <div align="left"><font class="form_dado_texto"> 
              Falecido</font></div></td>
          <td width="1%"><div align="center"><font class="form_dado_texto">?</font></div></td>
          <td width="15%" height="26"><font class="form_corpo"> 
                            <select name="pai_falecido" class="borda">
                              <option value="n"selected>N&atilde;o</option>
                              <option value="s">Sim</option>
                            </select>
            </font></td>
          <td width="15%" height="26"> <div align="left"><font class="form_dado_texto"> 
              Situa&ccedil;&atilde;o dos Pais</font></div></td>
          <td width="1%"><div align="center">:</div></td>
          <td width="15%" height="26"><font class="form_corpo"> 
<select name="sit_pais" class="borda" id="sit_pais">
                              <option value=0></option>
                              <%				
		Set RS_ec = Server.CreateObject("ADODB.Recordset")
		SQL_ec = "SELECT * FROM TB_Estado_Civil order by CO_Estado_Civil"
		RS_ec.Open SQL_ec, CON0
		
while not RS_ec.EOF						
co_ec= RS_ec("CO_Estado_Civil")
no_ec= RS_ec("TX_Estado_Civil")

if co_ec=sit_pais then
%>
                              <option value="<%=co_ec%>" selected> 
                              <% =no_ec%>
                              </option>
                              <%
else							  
%>
                              <option value="<%=co_ec%>"> 
                              <% =no_ec%>
                              </option>
                              <%
end if							  						
RS_ec.MOVENEXT
WEND
%>
                            </select>		  
            </font></td>
        </tr>
        <tr> 
          <td width="14%" height="10"> <div align="left"><font class="form_dado_texto"> 
              M&atilde;e</font></div></td>
          <td width="2%"><div align="center">: </div></td>
          <td height="10"><font class="form_corpo"> 
            <input name="mae" type="text" class="borda" onBlur="recuperarMae(this.value,'m','<%response.Write(mae_cadastrado)%>','<%response.Write(aluno)%>')" value="<%response.Write(mae)%>" size="30" maxlength="50">
            </font></td>
          <td height="10"> <div align="left"><font class="form_dado_texto"> Falecida</font></div></td>
          <td><div align="center"><font class="form_dado_texto">?</font></div></td>
          <td height="10"><font class="form_corpo"> 
                            <select name="mae_falecido" class="borda">
                              <option value="n"selected>N&atilde;o</option>
                              <option value="s">Sim</option>
                            </select>
            </font></td>
          <td height="10"><div align="left"><font class="form_dado_texto"> </font></div></td>
          <td><div align="center"></div></td>
          <td height="10"><font class="form_dado_texto">&nbsp; </font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td class="tb_tit">Familiares</td>
  </tr>
                  <tr class="tb_corpo"> 
                    <td> <div id="familiares"> </div></td>
                  </tr>
                  <tr class="tb_corpo">
                    <td><div id="responsaveis"> </div></td>
                  </tr>
</table>
<p><br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<br>
  <br>
  <br>
  <br>
  <br>
<%elseif opt="v" then

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& vinculado
		RS.Open SQL, CON1
		
			if RS.EOF Then
			else
			nome_prof = RS("NO_Aluno")
			sexo = RS("IN_Sexo")
			resp_fin= RS("TP_Resp_Fin")
			resp_ped= RS("TP_Resp_Ped")
			pai_fal= RS("IN_Pai_Falecido")
			mae_fal= RS("IN_Mae_Falecida")

			Set RS_aux = Server.CreateObject("ADODB.Recordset")
			SQL_aux = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="& vinculado
			RS_aux.Open SQL_aux, CON1_aux

			if RS_aux.EOF Then
				Set RSALUNO_aux_bd = server.createobject("adodb.recordset")
				RSALUNO_aux_bd.open "TBI_Alunos", CON1_aux, 2, 2
				RSALUNO_aux_bd.addnew
				RSALUNO_aux_bd("CO_Matricula")=vinculado
				RSALUNO_aux_bd("TP_Resp_Fin")=resp_fin							  
				RSALUNO_aux_bd("TP_Resp_Ped")=resp_ped
				RSALUNO_aux_bd("IN_Pai_Falecido")=pai_fal							  
				RSALUNO_aux_bd("IN_Mae_Falecida")=mae_fal				
				RSALUNO_aux_bd.update	
				set RSALUNO_aux_bd=nothing
			else
				Set RSALUNO_aux_bd2 = server.createobject("adodb.recordset")
				sql_atualiza_al= "UPDATE TBI_Alunos SET IN_Pai_Falecido="&pai_fal&", IN_Mae_Falecida="&mae_fal&", TP_Resp_Fin ='"& resp_fin &"', TP_Resp_Ped ='"& resp_ped &"' WHERE CO_Matricula = "& vinculado
				Set RSALUNO_aux_bd2 = CON1_aux.Execute(sql_atualiza_al)				
			end if

		end if	
'end if


		Set RSCONTA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT * FROM TB_Contatos WHERE TP_Contato ='ALUNO' And CO_Matricula ="& vinculado
		RSCONTA.Open SQLA, CONCONT

if RSCONTA.EOF then
nascimento="0/0/0"
else
nascimento = RSCONTA("DA_Nascimento_Contato")
end if
vetor_nascimento = Split(nascimento,"/")  
dia_n = vetor_nascimento(0)
mes_n = vetor_nascimento(1)
ano_n = vetor_nascimento(2)

if dia_n<10 then 
dia_n = "0"&dia_n
end if

if mes_n<10 then
mes_n = "0"&mes_n
end if
dia_a = dia_n
mes_a = mes_n
ano_a = ano_n

nasce = dia_n&"/"&mes_n&"/"&ano_n

apelido= RS("NO_Apelido")
desteridade= RS("IN_Desteridade")
nacionalidade= RS("CO_Nacionalidade")
rua_res = RSCONTA("NO_Logradouro_Res")
num_res = RSCONTA("NU_Logradouro_Res")
comp_res = RSCONTA("TX_Complemento_Logradouro_Res")
bairro_res= RSCONTA("CO_Bairro_Res")
cid_res= RSCONTA("CO_Municipio_Res")
pai= RS("NO_Pai")
mae= RS("NO_Mae")
pai_fal= RS("IN_Pai_Falecido")
mae_fal= RS("IN_Mae_Falecida")
uf_res= RSCONTA("SG_UF_Res")
cep = RSCONTA("CO_CEP_Res")
tel_res = RSCONTA("NU_Telefones_Res")
tel_cont = RSCONTA("NU_Telefones")
uf_natural = RS("SG_UF_Natural")
natural = RS("CO_Municipio_Natural")
resp_fin= RS("TP_Resp_Fin")
resp_ped= RS("TP_Resp_Ped")
mail= RSCONTA("TX_EMail")
pais= RS("CO_Pais_Natural")
ocupacao= RSCONTA("CO_Ocupacao")
msn= RS("TX_MSN")
orkut= RS("TX_ORKUT")
religiao= RS("CO_Religiao")
raca= RS("CO_Raca")
entrada= RS("DA_Entrada_Escola")
cadastro= RS("DA_Cadastro")
col_origem= RS("NO_Colegio_Origem")
cursada= RS("NO_Serie_Cursada")
uf_cursada= RS("SG_UF_Cursada")
cid_cursada= RS("CO_Municipio_Cursada")
sit_pais= RS("CO_Estado_Civil")
cpf= RSCONTA("CO_CPF_PFisica")
rg= RSCONTA("CO_RG_PFisica")
emitido= RSCONTA("CO_OERG_PFisica")
emissao= RSCONTA("CO_DERG_PFisica")
empresa= RSCONTA("NO_Empresa")
rua_com=RSCONTA("NO_Logradouro_Com")
num_com = RSCONTA("NU_Logradouro_Com")
comp_com = RSCONTA("TX_Complemento_Logradouro_Com")
bairro_com= RSCONTA("CO_Bairro_Com")
cid_com= RSCONTA("CO_Municipio_Com")
uf_com= RSCONTA("SG_UF_Com")
cep_com = RSCONTA("CO_CEP_Com")
tel_com = RSCONTA("NU_Telefones_Com")

			cpf=RSCONTA("CO_CPF_PFisica")
			rg=RSCONTA("CO_RG_PFisica")
			id_res_familiar=RSCONTA("ID_Res_Aluno")
			id_familia=RSCONTA("ID_Familia")
			id_end_bloq=RSCONTA("ID_End_Bloqueto")
			co_vinc_familiar=RSCONTA("CO_Matricula_Vinc")
			tp_vinc_familiar=RSCONTA("TP_Contato_Vinc")


'Grava dados de aluno Vinculado
		Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
		SQLAC_delete= "DELETE * FROM TBI_Contatos WHERE CO_Matricula ="&vinculado
		RSCONTATO_aux_delete.Open SQLAC_delete, CONCONT_aux
		
				Set RSCONTATO_aux_bd1 = server.createobject("adodb.recordset")
				RSCONTATO_aux_bd1.open "TBI_Contatos", CONCONT_aux, 2, 2 'which table do you want open
				RSCONTATO_aux_bd1.addnew
				RSCONTATO_aux_bd1("CO_Matricula")=vinculado
				RSCONTATO_aux_bd1("TP_Contato")="ALUNO"
				RSCONTATO_aux_bd1("NO_Contato")=nome_prof
				RSCONTATO_aux_bd1("DA_Nascimento_Contato")=nascimento
				RSCONTATO_aux_bd1("CO_CPF_PFisica")=cpf
				RSCONTATO_aux_bd1("CO_RG_PFisica")=rg
				RSCONTATO_aux_bd1("CO_OERG_PFisica")=emitido
				RSCONTATO_aux_bd1("CO_DERG_PFisica")=emissao
				RSCONTATO_aux_bd1("TX_EMail")=mail
				RSCONTATO_aux_bd1("CO_Ocupacao")=ocupacao
				RSCONTATO_aux_bd1("NO_Empresa")=empresa
				RSCONTATO_aux_bd1("NU_Telefones")=tel_cont
				RSCONTATO_aux_bd1("ID_Res_Aluno")=id_res_familiar
				RSCONTATO_aux_bd1("ID_Familia")=id_familia
				RSCONTATO_aux_bd1("ID_End_Bloqueto")=id_end_bloq
				RSCONTATO_aux_bd1("NO_Logradouro_Res")=rua_res
				RSCONTATO_aux_bd1("NU_Logradouro_Res")=num_res
				RSCONTATO_aux_bd1("TX_Complemento_Logradouro_Res")=comp_res
				RSCONTATO_aux_bd1("CO_Bairro_Res")=bairro_res
				RSCONTATO_aux_bd1("CO_Municipio_Res")=cid_res
				RSCONTATO_aux_bd1("SG_UF_Res")=uf_res
				RSCONTATO_aux_bd1("CO_CEP_Res")=cep_res
				RSCONTATO_aux_bd1("NU_Telefones_Res")=tel_res
				RSCONTATO_aux_bd1("NO_Logradouro_Com")=rua_com
				RSCONTATO_aux_bd1("NU_Logradouro_Com")=num_com
				RSCONTATO_aux_bd1("TX_Complemento_Logradouro_Com")=comp_com
				RSCONTATO_aux_bd1("CO_Bairro_Com")=bairro_com
				RSCONTATO_aux_bd1("CO_Municipio_Com")=cid_com
				RSCONTATO_aux_bd1("SG_UF_Com")=uf_com
				RSCONTATO_aux_bd1("CO_CEP_Com")=cep_com
				RSCONTATO_aux_bd1("NU_Telefones_Com")=tel_com
				RSCONTATO_aux_bd1("CO_Matricula_Vinc")=co_vinc_familiar
				RSCONTATO_aux_bd1("TP_Contato_Vinc")=tp_vinc_familiar
				RSCONTATO_aux_bd1.update
				set RSCONTATO_aux_bd1=nothing		



' Grava dados dos familiares vinculados ao aluno que está se vinculando
familiares = Split(ordem_familiares, "##")
	for i=1 to ubound(familiares)
		cod_nome_familiar=familiares(i)
		cod_nome = Split(cod_nome_familiar, "!!")
		cod_familiar=cod_nome(0)
		nome_familiar=cod_nome(1)		

		Set RSCONTATO = Server.CreateObject("ADODB.Recordset")
		SQLAA= "SELECT * FROM TB_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&vinculado
		RSCONTATO.Open SQLAA, CONCONT
		
		if RSCONTATO.EOF then
		else
			nome_familiar_aux=RSCONTATO("NO_Contato")
			nasce_familiar_aux=RSCONTATO("DA_Nascimento_Contato")
			cpf_familiar_aux=RSCONTATO("CO_CPF_PFisica")
			rg_familiar_aux=RSCONTATO("CO_RG_PFisica")
			emitido_familiar_aux=RSCONTATO("CO_OERG_PFisica")
			emissao_familiar_aux=RSCONTATO("CO_DERG_PFisica")
			email_familiar_aux=RSCONTATO("TX_EMail")
			ocupacao_familiar_aux=RSCONTATO("CO_Ocupacao")
			empresa_familiar_aux=RSCONTATO("NO_Empresa")
			tel_familiar_aux=RSCONTATO("NU_Telefones")
			id_res_familiar_aux=RSCONTATO("ID_Res_Aluno")
			id_familia_aux=RSCONTATO("ID_Familia")
			id_end_bloq_aux=RSCONTATO("ID_End_Bloqueto")
			rua_res_familiar_aux=RSCONTATO("NO_Logradouro_Res")
			num_res_familiar_aux=RSCONTATO("NU_Logradouro_Res")
			comp_res_familiar_aux=RSCONTATO("TX_Complemento_Logradouro_Res")
			bairro_res_familiar_aux=RSCONTATO("CO_Bairro_Res")
			cid_res_familiar_aux=RSCONTATO("CO_Municipio_Res")
			uf_res_familiar_aux=RSCONTATO("SG_UF_Res")
			cep_res_familiar_aux=RSCONTATO("CO_CEP_Res")
			tel_res_familiar_aux=RSCONTATO("NU_Telefones_Res")
			rua_com_familiar_aux=RSCONTATO("NO_Logradouro_Com")
			num_com_familiar_aux=RSCONTATO("NU_Logradouro_Com")
			comp_com_familiar_aux=RSCONTATO("TX_Complemento_Logradouro_Com")
			bairro_com_familiar_aux=RSCONTATO("CO_Bairro_Com")
			cid_com_familiar_aux=RSCONTATO("CO_Municipio_Com")
			uf_com_familiar_aux=RSCONTATO("SG_UF_Com")
			cep_com_familiar_aux=RSCONTATO("CO_CEP_Com")
			tel_com_familiar_aux=RSCONTATO("NU_Telefones_Com")
			co_vinc_familiar_aux=RSCONTATO("CO_Matricula_Vinc")
			tp_vinc_familiar_aux=RSCONTATO("TP_Contato_Vinc")

			Set RSCONTATO_aux = Server.CreateObject("ADODB.Recordset")
			SQLAA_aux= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&vinculado
			RSCONTATO_aux.Open SQLAA_aux, CONCONT_aux

			if RSCONTATO_aux.EOF then
				Set RSCONTATO_aux_bd = server.createobject("adodb.recordset")

				RSCONTATO_aux_bd.open "TBI_Contatos", CONCONT_aux, 2, 2 'which table do you want open
				RSCONTATO_aux_bd.addnew
				RSCONTATO_aux_bd("CO_Matricula")=vinculado
				RSCONTATO_aux_bd("TP_Contato")=cod_familiar
				RSCONTATO_aux_bd("NO_Contato")=nome_familiar_aux
				RSCONTATO_aux_bd("DA_Nascimento_Contato")=nasce_familiar_aux
				RSCONTATO_aux_bd("CO_CPF_PFisica")=cpf_familiar_aux
				RSCONTATO_aux_bd("CO_RG_PFisica")=rg_familiar_aux
				RSCONTATO_aux_bd("CO_OERG_PFisica")=emitido_familiar_aux
				RSCONTATO_aux_bd("CO_DERG_PFisica")=emissao_familiar_aux
				RSCONTATO_aux_bd("TX_EMail")=email_familiar_aux
				RSCONTATO_aux_bd("CO_Ocupacao")=ocupacao_familiar_aux
				RSCONTATO_aux_bd("NO_Empresa")=empresa_familiar_aux
				RSCONTATO_aux_bd("NU_Telefones")=tel_familiar_aux
				RSCONTATO_aux_bd("ID_Res_Aluno")=id_res_familiar_aux
				RSCONTATO_aux_bd("ID_Familia")=id_familia_aux
				RSCONTATO_aux_bd("ID_End_Bloqueto")=id_end_bloq_aux
				RSCONTATO_aux_bd("NO_Logradouro_Res")=rua_res_familiar_aux
				RSCONTATO_aux_bd("NU_Logradouro_Res")=num_res_familiar_aux
				RSCONTATO_aux_bd("TX_Complemento_Logradouro_Res")=comp_res_familiar_aux
				RSCONTATO_aux_bd("CO_Bairro_Res")=bairro_res_familiar_aux
				RSCONTATO_aux_bd("CO_Municipio_Res")=cid_res_familiar_aux
				RSCONTATO_aux_bd("SG_UF_Res")=uf_res_familiar_aux
				RSCONTATO_aux_bd("CO_CEP_Res")=cep_res_familiar_aux
				RSCONTATO_aux_bd("NU_Telefones_Res")=tel_res_familiar_aux
				RSCONTATO_aux_bd("NO_Logradouro_Com")=rua_com_familiar_aux
				RSCONTATO_aux_bd("NU_Logradouro_Com")=num_com_familiar_aux
				RSCONTATO_aux_bd("TX_Complemento_Logradouro_Com")=comp_com_familiar_aux
				RSCONTATO_aux_bd("CO_Bairro_Com")=bairro_com_familiar_aux
				RSCONTATO_aux_bd("CO_Municipio_Com")=cid_com_familiar_aux
				RSCONTATO_aux_bd("SG_UF_Com")=uf_com_familiar_aux
				RSCONTATO_aux_bd("CO_CEP_Com")=cep_com_familiar_aux
				RSCONTATO_aux_bd("NU_Telefones_Com")=tel_com_familiar_aux
				RSCONTATO_aux_bd("CO_Matricula_Vinc")=co_vinc_familiar_aux
				RSCONTATO_aux_bd("TP_Contato_Vinc")=tp_vinc_familiar_aux
				RSCONTATO_aux_bd.update
				set RSCONTATO_aux_bd=nothing
				

				
		Set RSCONTATO_aux_delete = Server.CreateObject("ADODB.Recordset")
		SQLAC_delete= "DELETE * FROM TBI_Contatos WHERE TP_Contato = '"&cod_familiar&"' and CO_Matricula ="&aluno
		RSCONTATO_aux_delete.Open SQLAC_delete, CONCONT_aux

				Set RSCONTATO_aux_bd_vincula = server.createobject("adodb.recordset")				
				RSCONTATO_aux_bd_vincula.open "TBI_Contatos", CONCONT_aux, 2, 2 'which table do you want open
				RSCONTATO_aux_bd_vincula.addnew
				RSCONTATO_aux_bd_vincula("CO_Matricula")=aluno
				RSCONTATO_aux_bd_vincula("TP_Contato")=cod_familiar
				if vinculado<>aluno then
					RSCONTATO_aux_bd_vincula("CO_Matricula_Vinc")=vinculado
					RSCONTATO_aux_bd_vincula("TP_Contato_Vinc")=cod_familiar
				end if				
				RSCONTATO_aux_bd_vincula.update
				set RSCONTATO_aux_bd_vincula=nothing

			else

				if isnull(nasce_familiar_aux) or nasce_familiar_aux="" then
				sql_nasce="DA_Nascimento_Contato =NULL"
				else
				sql_nasce="DA_Nascimento_Contato =#"& nasce_familiar_aux &"#"
				end if

				if isnull(emissao_familiar_aux) or emissao_familiar_aux="" then
				sql_emissao="CO_DERG_PFisica =NULL"
				else
				sql_emissao="CO_DERG_PFisica =#"& emissao_familiar_aux &"#"
				end if

				if isnull(ocupacao_familiar_aux) or ocupacao_familiar_aux="" then
				sql_ocupacao="CO_Ocupacao =NULL"
				else
				sql_ocupacao="CO_Ocupacao ="& ocupacao_familiar_aux &""
				end if

				if isnull(num_res_familiar_aux) or num_res_familiar_aux="" then
				sql_num_res="NU_Logradouro_Res =NULL"
				else
				sql_num_res="NU_Logradouro_Res ="& num_res_familiar_aux &""
				end if

				if isnull(bairro_res_familiar_aux) or bairro_res_familiar_aux="" then
				sql_bairro_res=" CO_Bairro_Res =NULL"
				else
				sql_bairro_res=" CO_Bairro_Res ="& bairro_res_familiar_aux &""
				end if

				if isnull(cid_res_familiar_aux) or cid_res_familiar_aux="" then
				sql_cid_res=" CO_Municipio_Res =NULL"
				else
				sql_cid_res=" CO_Municipio_Res ="& cid_res_familiar_aux &""
				end if

				if isnull(num_com_familiar_aux) or num_com_familiar_aux="" then
				sql_num_com="NU_Logradouro_Com =NULL"
				else
				sql_num_com="NU_Logradouro_Com ="& num_com_familiar_aux &""
				end if

				if isnull(cid_com_familiar_aux) or cid_com_familiar_aux="" then
				sql_cid_com=" CO_Municipio_Com =NULL"
				else
				sql_cid_com=" CO_Municipio_Com ="& cid_com_familiar_aux &""
				end if

				if isnull(bairro_com_familiar_aux) or bairro_com_familiar_aux="" then
				sql_bairro_com=" CO_Bairro_Com =NULL"
				else
				sql_bairro_com=" CO_Bairro_Com ="& bairro_com_familiar_aux &""
				end if

				if isnull(co_vinc_familiar_aux) or co_vinc_familiar_aux="" then
				sql_vinc="CO_Matricula_Vinc =NULL"
				else
				sql_vinc="CO_Matricula_Vinc ="& co_vinc_familiar_aux &""
				end if
				
				Set RSCONTATO_aux_bd2 = server.createobject("adodb.recordset")
				sql_atualiza= "UPDATE TBI_Contatos SET NO_Contato = '"&nome_familiar_aux&"', "& sql_nasce &", CO_CPF_PFisica ='"& cpf_familiar_aux &"', CO_RG_PFisica ='"& rg_familiar_aux &"', CO_OERG_PFisica ='"& emitido_familiar_aux&"', "& sql_emissao &", TX_EMail ='"& email_familiar_aux &"', "&sql_ocupacao&", NO_Empresa ='"& empresa_familiar_aux &"', "
				sql_atualiza=sql_atualiza&"NU_Telefones ='"& tel_familiar_aux&"', ID_Res_Aluno = "&id_res_familiar_aux&", NO_Logradouro_Res ='"& rua_res_familiar_aux &"', "& sql_num_res&", TX_Complemento_Logradouro_Res = '"&comp_res_familiar_aux&"', "& sql_bairro_res &", "& sql_cid_res &", SG_UF_Res ='"& uf_res_familiar_aux &"', "
				sql_atualiza=sql_atualiza&"CO_CEP_Res ='"& cep_res_familiar_aux &"', NU_Telefones_Res ='"& tel_res_familiar_aux&"', NO_Logradouro_Com = '"&rua_com_familiar_aux&"', "& sql_num_com&", TX_Complemento_Logradouro_Com= '"&comp_com_familiar_aux&"',"&sql_bairro_com&", "& sql_cid_com &", SG_UF_Com ='"& uf_com_familiar_aux &"', "
				sql_atualiza=sql_atualiza&"CO_CEP_Com='"& cep_com_familiar_aux &"', NU_Telefones_Com ='"& tel_com_familiar_aux&"', "&sql_vinc&", TP_Contato_Vinc ='"& tp_vinc_familiar_aux&"' WHERE CO_Matricula = "& vinculado &" AND TP_Contato = '"& cod_familiar &"'"
				Set RSCONTATO_aux_bd2 = CONCONT_aux.Execute(sql_atualiza)

				Set RSCONTATO_aux_bd2_vincula = server.createobject("adodb.recordset")
				if vinculado<>aluno then
					sql_atualiza_vincula= "UPDATE TBI_Contatos SET CO_Matricula_Vinc ="& vinculado &", TP_Contato_Vinc ='"& cod_familiar&"' WHERE CO_Matricula = "& aluno &" AND TP_Contato = '"& cod_familiar &"'"
				else
					sql_atualiza_vincula= "UPDATE TBI_Contatos SET CO_Matricula_Vinc =NULL, TP_Contato_Vinc ='' WHERE CO_Matricula = "& aluno &" AND TP_Contato = '"& cod_familiar &"'"
				end if
				Set RSCONTATO_aux_bd2_vincula = CONCONT_aux.Execute(sql_atualiza_vincula)
				
			end if
		end if
	next
'Grava vínculo ao aluno	
				Set RSALUNO_vincula = server.createobject("adodb.recordset")
				if vinculado<>aluno then
					sql_atualiza_vincula= "UPDATE TBI_Contatos SET CO_Matricula_Vinc ="& vinculado &", TP_Contato_Vinc ='ALUNO' WHERE CO_Matricula = "& aluno &" AND TP_Contato = 'ALUNO'"
				else
					sql_atualiza_vincula= "UPDATE TBI_Contatos SET CO_Matricula_Vinc =NULL, TP_Contato_Vinc ='' WHERE CO_Matricula = "& aluno &" AND TP_Contato = 'ALUNO'"
				end if
				Set RSALUNO_vincula = CONCONT_aux.Execute(sql_atualiza_vincula)

if isnull(pais) then 
pais = 10
end if

if isnull(uf) then 
uf = "RJ"
end if

if isnull(municipio) then 
municipio = 6001
end if

if isnull(uf_natural) then 
uf_natural = "RJ"
end if

if isnull(nacionalidade) then 
nacionalidade = 1
end if

if isnull(natural) then 
natural = 6001
end if

if complemento = "nulo" then 
complemento = ""
end if


if desteridade = "S" then
desteridade = "Destro"
else
desteridade = "Canhoto"
end if

if isnull(cid_cursada) then 
cid_cursada = 6001
end if

if isnull(uf_cursada) then 
uf_cursada = "RJ"
end if

if isnull(cep) or cep="" then
cep=""
else
cep5= lEFT(cep, 5)
cep3= Right(cep, 3)
cep=cep5&"-"&cep3
end if

if isnull(cepcom) or cepcom="" then
cep2=""
else
cep5c= lEFT(cepcom, 5)
cep3c= Right(cepcom, 3)
cepcom=cep5c&"-"&cep3c
end if

		Set RS0 = Server.CreateObject("ADODB.Recordset")
		SQL0 = "SELECT * FROM TB_Religiao WHERE CO_Religiao ="& religiao
		RS0.Open SQL0, CON0

religiao = RS0("TX_Descricao_Religiao")

		Set RS1 = Server.CreateObject("ADODB.Recordset")
		SQL1 = "SELECT * FROM TB_Raca WHERE CO_Raca ="& raca
		RS1.Open SQL1, CON0

raca = RS1("TX_Descricao_Raca")

		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_Paises WHERE CO_Pais ="& pais
		RS2.Open SQL2, CON0

pais = RS2("NO_Pais")

		Set RS3 = Server.CreateObject("ADODB.Recordset")
		SQL3 = "SELECT * FROM TB_Nacionalidades WHERE CO_Nacionalidade ="& nacionalidade
		RS3.Open SQL3, CON0

nacionalidade = RS3("TX_Nacionalidade")



		Set RS6 = Server.CreateObject("ADODB.Recordset")
		SQL6 = "SELECT * FROM TB_Municipios WHERE SG_UF ='"& uf_natural&"' AND CO_Municipio = "&natural
		RS6.Open SQL6, CON0

natural= RS6("NO_Municipio")

		Set RS8 = Server.CreateObject("ADODB.Recordset")
		SQL8 = "SELECT * FROM TB_UF WHERE SG_UF ='"& uf_natural&"'" 
		RS8.Open SQL8, CON0

uf_natural= RS8("NO_UF")

		Set RS9 = Server.CreateObject("ADODB.Recordset")
		SQL9 = "SELECT * FROM TB_Ocupacoes WHERE CO_Ocupacao ="& ocupacao
		RS9.Open SQL9, CON0

ocupacao= RS9("NO_Ocupacao")

		Set RSCONTPR = Server.CreateObject("ADODB.Recordset")
		SQLCONTPR = "SELECT * FROM TB_Tipo_Contatos Order by NU_Prioridade_Combo"
		RSCONTPR.Open SQLCONTPR, CON0
total_tp_familiares=0
pai_cadastrado="n"
mae_cadastrado="n"

while not RSCONTPR.EOF
tp_resp = RSCONTPR("TP_Contato")

		Set RSCONTACONT = Server.CreateObject("ADODB.Recordset")
		SQLAC = "SELECT * FROM TBI_Contatos WHERE TP_Contato='"& tp_resp&"' AND CO_Matricula ="& aluno
		RSCONTACONT.Open SQLAC, CONCONT_aux
familiares=0
if RSCONTACONT.EOF then
foco=foco_default
else
	while not RSCONTACONT.EOF
	tipo_contato=RSCONTACONT("TP_Contato")
		if familiares=0 then
		foco=tipo_contato
		end if
		
		if tipo_contato="PAI" OR tipo_contato="MAE" OR tipo_contato="ALUNO" then
				if tipo_contato="PAI" then
				pai_cadastrado="s"
				end if				

				if tipo_contato="MAE" then
				mae_cadastrado="s"
				end if				

				RSCONTACONT.MOVENEXT
		else
		familiares=familiares+1
		nome_contato=RSCONTACONT("NO_Contato")
		RSCONTACONT.MOVENEXT
		end if
	wend
end if		
RSCONTPR.MOVENEXT
WEND		
%>


<table width="1000" height="100" border="0" align="right" cellspacing="0" class="tb_corpo">
  <td class="tb_tit"
>Dados vinculados ao Aluno</td>
          </tr>
          <tr> 
            <td> 
<% 
aluno=aluno*1
vinculado=vinculado*1
if aluno=vinculado then

no_alt_pai_mae="s"
%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr bgcolor="#FFFFFF" background="../../../../img/fundo_interno.gif"> 
                          <td width="5%"  height="10"> <div align="left"><font class="form_dado_texto"> 
                              Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                              </strong></font></div></td>
                          <td width="10%" height="10"><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                            </font><font size="2" face="Arial, Helvetica, sans-serif"> 
                            <input name="vinculado_novo" type="text" class="borda" id="vinculado" size="12">
                            </font><font size="1" face="Verdana, Arial, Helvetica, sans-serif">&nbsp; 
                            </font></td>
                          <td width="37%" height="10"> <div align="right"><font class="form_dado_texto"> 
                              </font></div></td>
                          <td width="2%" height="10" ><font size="2" face="Arial, Helvetica, sans-serif">&nbsp; 
                            </font></td>
                          <td width="46%" height="10"><font size="2" face="Arial, Helvetica, sans-serif">  <!--  'O valor de codigo.value vem do arquivo altera.asp-->
                            <input name="Button" type="button" class="borda_bot" id="Submit" value="Vincular" onClick="VincularAluno(vinculado_novo.value,codigo.value)">
                            </font> </td>
                        </tr>
                      </table>
<%else
no_alt_pai_mae="n"
%>			
			<table width="100%" border="0" cellspacing="0" cellpadding="0">
                <tr bgcolor="#FFFFFF" background="../../../../img/fundo_interno.gif"> 
                          <td width="5%"  height="10"> <div align="left"><font class="form_dado_texto"> 
                              Matr&iacute;cula:</font><font size="1" face="Verdana, Arial, Helvetica, sans-serif"><strong> 
                              </strong></font></div></td>
                  <td width="10%" height="10"><div align="left"><font class="form_dado_texto"> 
              <%response.Write(Server.URLEncode(vinculado))%>
              <input name="vinculado" type="hidden" class="borda" id="vinculado" value="<%response.Write(vinculado)%>" size="12">
              <input name="aluno" type="hidden" class="borda" id="aluno" value="<%response.Write(aluno)%>" size="12">
              </font></div>
                    </td>
                  
          <td width="37%" height="10"> <div align="left"><font class="form_dado_texto"> 
              <%response.Write(Server.URLEncode(nome_prof))%>
              </font></div></td>
                  <td width="2%" height="10" ><font size="2" face="Arial, Helvetica, sans-serif">&nbsp; 
                    </font></td>
                  <td width="46%" height="10"><font size="2" face="Arial, Helvetica, sans-serif">
                    <input type="button" name="Button" value="Desvincular"  class="borda_bot3" onClick="DesvincularAluno(vinculado.value,aluno.value)">
                    </font> </td>
                </tr>
              </table>
<%end if%>			  
			  </td>
          </tr>			
  <tr> 
    <td class="tb_tit"
>Endere&ccedil;o Residencial</td>
  </tr>
  <tr> 
    <td height="10"><table width="100%" border="0" cellspacing="0">
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                          <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                          <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                            
            <input name="rua_res" type="text" class="borda" id="rua_res" value="<%response.Write(rua_res)%>" size="30" maxlength="60">
                            </font></td>
                          <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                          <td width="19" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="196" class="tb_corpo"
><font class="form_corpo"> <font class="form_corpo"> <font class="form_corpo"> 
                            <input name="num_res" type="text" class="borda" id="num_res" value="<%response.Write(num_res)%>" size="12" maxlength="10" onBlur="ValidaNumResFam(this.value)">
                            </font></font></font></td>
                          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                          <td width="11" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="149" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
                              
              <input name="comp_res" type="text" class="borda" id="comp_res" value="<%response.Write(comp_res)%>" size="20" maxlength="30">
                              </font></div></td>
                        </tr>
                        <tr> 
                          <td width="145" height="21" class="tb_corpo"
><font class="form_dado_texto">Estado</font></td>
                          <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                          <td width="217" height="21" class="tb_corpo"
><font class="form_corpo"> <font class="form_corpo"> 
                            <select name="uf_res" class="borda" id="select4" onChange="recuperarCidRes(this.value)">
                              <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")

if isnull(uf_res) or uf_res="" then
uf_res="RJ"
end if

if SG_UF = uf_res then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <%response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <% else %>
                                <option value="<%=SG_UF%>"> 
                                <% response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                              <%
end if	
RS2.MOVENEXT
WEND
%>
                            </select>
                            </font> </font></td>
                          <td width="140" height="21" class="tb_corpo"
><font class="form_dado_texto">Cidade</font></td>
                          <td width="19" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                            <div id="cid_res"> 
                              <select name="cid_res" class="borda" id="select10" onChange="recuperarBairroRes(estadores.value,this.value)">
                                <%
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='"&uf_res&"' order by NO_Municipio"
		RS2m.Open SQL2m, CON0
		
while not RS2m.EOF						
SG_UF= RS2m("CO_Municipio")
NO_UF= RS2m("NO_Municipio")
if isnull(cid_res) or cid_res="" then
cid_res=6001
end if
if SG_UF = cid_res then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <%response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <% else %>
                                <option value="<%=SG_UF%>"> 
                                <% response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <%
end if	
RS2m.MOVENEXT
WEND
%>
                              </select>
                            </div>
                            </font></td>
                          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Bairro</font></td>
                          <td width="11" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="149" height="21" class="tb_corpo"
> <div id="bairro_res"><font class="form_corpo"> 
                              <select name="bairro_res" class="borda" id="select3">
                                <%
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Municipio="&cid_res&" AND SG_UF='"&uf_res&"' order by NO_Bairro"
		RS2b.Open SQL2b, CON0
		
while not RS2b.EOF						
SG_UF= RS2b("CO_Bairro")
NO_UF= RS2b("NO_Bairro")
if SG_UF=bairro_res then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <%response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <% else %>
                                <option value="<%=SG_UF%>"> 
                                <% response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <%
end if
RS2b.MOVENEXT
WEND
%>
                              </select>
                              </font></div></td>
                        </tr>
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
><font class="form_dado_texto">CEP</font></td>
                          <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                          <td width="217" height="10" class="tb_corpo"
><font class="form_dado_texto"> 
                            <input name="cep" type="text" class="borda" id="cep" onKeyup="formatar(this, '#####-###')" value="<%response.Write(cep)%>" size="11" maxlength="9" onBlur="ValidaCepResFam(this.value)">
                            </font></td>
                          <td width="140" height="10" class="tb_corpo"
>&nbsp;</td>
                          <td width="19" class="tb_corpo"
>&nbsp;</td>
                          <td width="196" class="tb_corpo"
>&nbsp;</td>
                          <td width="90" class="tb_corpo"
>&nbsp;</td>
                          <td width="11" class="tb_corpo"
> <div align="center"></div></td>
                          <td width="149" height="10" class="tb_corpo"
>&nbsp;</td>
                        </tr>
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">Telefones deste endere&ccedil;o</font></div></td>
                          <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                          <td height="10" colspan="2" class="tb_corpo"
><font class="form_corpo"> 
                            
            <input name="tel_res" type="text" class="borda" id="tel_res" value="<%response.Write(tel_res)%>" size="42" maxlength="100">
                            </font> <div align="left"></div></td>
                          <td width="19" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="196" class="tb_corpo"
>&nbsp;</td>
                          <td width="90" class="tb_corpo"
>&nbsp;</td>
                          <td width="11" class="tb_corpo"
> <div align="center"></div></td>
                          <td width="149" height="10" class="tb_corpo"
>&nbsp;</td>
                        </tr>
                        <tr> 
                          <td height="10" colspan="9" class="tb_tit"
><div align="left">Endere&ccedil;o Comercial </div></td>
                        </tr>
                        <tr> 
                          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                          <td width="13" class="tb_corpo"
> <div align="left">:</div></td>
                          <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                            
            <input name="rua_com" type="text" class="borda" id="rua_com" value="<%response.Write(rua_com)%>" size="30" maxlength="60">
                            </font></td>
                          <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                          <td width="19" class="tb_corpo"
> <div align="center">:</div></td>
                          <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                            <input name="num_com" type="text" class="borda" id="num_com" value="<%response.Write(num_com)%>" size="12" maxlength="10" onBlur="ValidaNumComFam(this.value)">
                            </font></td>
                          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                          <td class="tb_corpo"
><div align="center">:</div></td>
                          <td width="149" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
                              
              <input name="comp_com" type="text" class="borda" id="comp_com"  value="<%response.Write(comp_com)%>" size="20" maxlength="30">
                              </font></div></td>
                        </tr>
                        <tr class="tb_corpo"
> 
                          <td width="145" height="26"><font class="form_dado_texto">Estado</font></td>
                          <td width="13"> <div align="left">:</div></td>
                          <td width="217" height="26"><font class="form_corpo"> 
                            <font class="form_corpo"> 
                            <select name="uf_com" class="borda" id="select2" onChange="recuperarCidCom(this.value)">
                              <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")
if isnull(uf_com) or uf_com=""  then
uf_com="RJ"
end if
if SG_UF = uf_com then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <%response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <% else %>
                                <option value="<%=SG_UF%>"> 
                                <% response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                              <%end if						
RS2.MOVENEXT
WEND
%>
                            </select>
                            </font> </font></td>
                          <td width="140" height="26"><font class="form_dado_texto">Cidade</font></td>
                          <td width="19"> <div align="center">:</div></td>
                          <td width="196"> <div id="cid_com"> 
                              <select name="cid_com" class="borda" id="select10" onChange="recuperarBairroCom(estadocom.value,this.value)">
                                <%
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='"&uf_com&"' order by NO_Municipio"
		RS2m.Open SQL2m, CON0
		
while not RS2m.EOF						
SG_UF= RS2m("CO_Municipio")
NO_UF= RS2m("NO_Municipio")
if isnull(cid_com) or cid_com="" then
cid_com=6001
end if
if SG_UF = cid_com then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <%response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <% else %>
                                <option value="<%=SG_UF%>"> 
                                <% response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <%
end if	
RS2m.MOVENEXT
WEND
%>
                              </select>
                            </div></td>
                          <td width="90"><font class="form_dado_texto">Bairro</font></td>
                          <td><div align="center">:</div></td>
                          <td width="149" height="26"> <div id="bairro_com"><font class="form_corpo"> 
                              <select name="bairro_com" class="borda" id="bairro">
                                <%
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Municipio="&cid_com&" AND SG_UF='"&uf_com&"' order by NO_Bairro"
		RS2b.Open SQL2b, CON0
		
while not RS2b.EOF						
SG_UF= RS2b("CO_Bairro")
NO_UF= RS2b("NO_Bairro")

if SG_UF=bairro_com then
%>
                                <option value="<%=SG_UF%>" selected> 
                                <%response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <% else %>
                                <option value="<%=SG_UF%>"> 
                                <% response.Write(Server.URLEncode(NO_UF))%>
                                </option>
                                <%
end if

RS2b.MOVENEXT
WEND
%>
                              </select>
                              </font></div></td>
                        </tr>
                        <tr class="tb_corpo"
> 
                          <td width="145" height="26"><font class="form_dado_texto">CEP</font></td>
                          <td width="13"> <div align="left">:</div></td>
                          <td width="217" height="26"><font class="form_dado_texto"> 
                            <input name="cep_com" type="text" class="borda" id="cepcom" onKeyup="formatar(this, '#####-###')" value="<%response.Write(cep_com)%>" size="11" maxlength="9" onBlur="ValidaCepComFam(this.value)">
                            </font></td>
                          <td width="140" height="26">&nbsp;</td>
                          <td width="19">&nbsp;</td>
                          <td width="196">&nbsp;</td>
                          <td width="90">&nbsp;</td>
                          <td><div align="center"></div></td>
                          <td width="149" height="26">&nbsp;</td>
                        </tr>
                        <tr class="tb_corpo"
> 
                          <td width="145" height="28"> <div align="left"><font class="form_dado_texto">Telefones 
                              deste endere&ccedil;o:</font></div></td>
                          <td width="13"> <div align="left">:</div></td>
                          <td height="28" colspan="2"><font class="form_corpo"> 
                            
            <input name="tel_com" type="text" class="borda" id="tel_com" value="<%response.Write(tel_com)%>" size="42" maxlength="100">
                            </font> <div align="left"></div></td>
                          <td width="19"> <div align="center"></div></td>
                          <td width="196">		  <marquee id="mqLooper1" loop="1"  onStart="<%response.Write("recuperarFamiliares('"&Server.URLEncode(ordem_familiares)&"','"&total_tp_familiares&"','"&foco&"','"&vinculado&"','"&aluno&"')")%>"></marquee>
</td>
                          <td width="90">&nbsp;</td>
                          <td><div align="center"></div></td>
                          <td width="149" height="28">&nbsp;</td>
                        </tr>
                      </table></td>
  </tr>
  <tr> 
    <td class="tb_tit"
>Filia&ccedil;&atilde;o</td>
  </tr>
  <tr> 
    <td><table width="100%" border="0" cellspacing="0">
        <tr> 
          <td width="14%" height="26"> <div align="left"><font class="form_dado_texto"> 
              Pai</font></div></td>
          <td width="2%"><div align="center">:</div></td>
          <td width="22%" height="26"><font class="form_corpo"> 
            <%if no_alt_pai_mae="s" then%>
            <input name="pai" type="text" class="borda" value="<%response.Write(pai)%>" size="30" maxlength="50">
            <%else%>
            <input name="pai" type="text" class="borda" onBlur="recuperarPai(this.value,'p','<%response.Write(pai_cadastrado)%>','<%response.Write(aluno)%>')" value="<%response.Write(pai)%>" size="30" maxlength="50">
            <%end if%>
            </font></td>
          <td width="15%" height="26"> <div align="left"><font class="form_dado_texto"> 
              Falecido</font></div></td>
          <td width="1%"><div align="center"><font class="form_dado_texto">?</font></div></td>
          <td width="15%" height="26"><font class="form_corpo"> 
                            <select name="pai_falecido" class="borda">
                              <% if pai_fal = false then%>
                              <option value="n"selected>N&atilde;o</option>
                              <option value="s">Sim</option>
                              <%else%>
                              <option value="n">N&atilde;o</option>
                              <option value="s" selected>Sim</option>
                              <%end if%>
                            </select>
            </font></td>
          <td width="15%" height="26"> <div align="left"><font class="form_dado_texto"> 
              Situa&ccedil;&atilde;o dos Pais</font></div></td>
          <td width="1%"><div align="center">:</div></td>
          <td width="15%" height="26"><font class="form_corpo"> 
<select name="sit_pais" class="borda" id="sit_pais">
                              <option value=0></option>
                              <%				
		Set RS_ec = Server.CreateObject("ADODB.Recordset")
		SQL_ec = "SELECT * FROM TB_Estado_Civil order by CO_Estado_Civil"
		RS_ec.Open SQL_ec, CON0
		
while not RS_ec.EOF						
co_ec= RS_ec("CO_Estado_Civil")
no_ec= RS_ec("TX_Estado_Civil")

if co_ec=sit_pais then
%>
                              <option value="<%=co_ec%>" selected> 
                              <% =no_ec%>
                              </option>
                              <%
else							  
%>
                              <option value="<%=co_ec%>"> 
                              <% =no_ec%>
                              </option>
                              <%
end if							  						
RS_ec.MOVENEXT
WEND
%>
                            </select>		  
            </font></td>
        </tr>
        <tr> 
          <td width="14%" height="10"> <div align="left"><font class="form_dado_texto"> 
              M&atilde;e</font></div></td>
          <td width="2%"><div align="center">: </div></td>
          <td height="10"><font class="form_corpo"> 
            <%if no_alt_pai_mae="s" then%>
            <input name="mae" type="text" class="borda" value="<%response.Write(mae)%>" size="30" maxlength="50">
            <%else%>		  
            <input name="mae" type="text" class="borda" onBlur="recuperarMae(this.value,'m','<%response.Write(mae_cadastrado)%>','<%response.Write(aluno)%>')" value="<%response.Write(mae)%>" size="30" maxlength="50">
			<%end if%>           
		    </font></td>
          <td height="10"> <div align="left"><font class="form_dado_texto"> Falecida</font></div></td>
          <td><div align="center"><font class="form_dado_texto">?</font></div></td>
          <td height="10"><font class="form_corpo"> 
                            <select name="mae_falecido" class="borda">
                              <% if mae_fal = false then%>
                              <option value="n"selected>N&atilde;o</option>
                              <option value="s">Sim</option>
                              <%else%>
                              <option value="n">N&atilde;o</option>
                              <option value="s" selected>Sim</option>
                              <%end if%>
                            </select>
            </font></td>
          <td height="10"><div align="left"><font class="form_dado_texto"> </font></div></td>
          <td><div align="center"></div></td>
          <td height="10"><font class="form_dado_texto">&nbsp; </font></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td class="tb_tit">Familiares</td>
  </tr>
                  <tr class="tb_corpo"> 
                    <td> <div id="familiares"> </div></td>
                  </tr>
                  <tr class="tb_corpo">
                    <td><div id="responsaveis"> </div></td>
                  </tr>
</table>
<p><br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
  <br>
</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<p>&nbsp;</p>
<br>
  <br>
  <br>
  <br>
  <br>
  <%
'recuperar endereço  
elseif opt="r" then
cod= request.form("cod_pub")
foco= request.form("foco_pub")	

		Set RS = Server.CreateObject("ADODB.Recordset")
		SQL = "SELECT * FROM TB_Alunos WHERE CO_Matricula ="& cod
		RS.Open SQL, CON1
		
		
codigo = RS("CO_Matricula")
nome_aluno = RS("NO_Aluno")

sexo = RS("IN_Sexo")

		Set RSCONTA = Server.CreateObject("ADODB.Recordset")
		SQLA = "SELECT * FROM TB_Contatos WHERE TP_Contato ='ALUNO' And CO_Matricula ="& cod
		RSCONTA.Open SQLA, CONCONT

rua_res = RSCONTA("NO_Logradouro_Res")
num_res = RSCONTA("NU_Logradouro_Res")
comp_res = RSCONTA("TX_Complemento_Logradouro_Res")
bairrores= RSCONTA("CO_Bairro_Res")
cidres= RSCONTA("CO_Municipio_Res")
estadores= RSCONTA("SG_UF_Res")
cep = RSCONTA("CO_CEP_Res")
tel_res = RSCONTA("NU_Telefones_Res")
tel = RSCONTA("NU_Telefones")
empresa= RSCONTA("NO_Empresa")
rua_com=RSCONTA("NO_Logradouro_Com")
num_com = RSCONTA("NU_Logradouro_Com")
comp_com = RSCONTA("TX_Complemento_Logradouro_Com")
bairrocom= RSCONTA("CO_Bairro_Com")
cidcom= RSCONTA("CO_Municipio_Com")
estadocom= RSCONTA("SG_UF_Com")
cepcom = RSCONTA("CO_CEP_Com")
tel_com = RSCONTA("NU_Telefones_Com")

session("id_res_familiar")="s"


if isnull(pais) then 
pais = 10
end if

if isnull(estadores) then 
estadores = "RJ"
end if

if isnull(cidres) then 
cidres = 6001
end if

if isnull(estadonat) then 
estadonat = "RJ"
end if

if isnull(nacionalidade) then 
nacionalidade = 1
end if

if isnull(cidnat) then 
cidnat = 6001
end if

if comp_res = "nulo" then 
comp_res = ""
end if

if isnull(cid_cursada) then 
cid_cursada = 6001
end if

if isnull(uf_cursada) then 
uf_cursada = "RJ"
end if


cep5= lEFT(cep, 5)
cep3= Right(cep, 3)


cep=cep5&"-"&cep3

cep5c= lEFT(cepcom, 5)
cep3c= Right(cepcom, 3)


cepcom=cep5c&"-"&cep3c


%>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr class="tb_corpo"> 
                          <td class="tb_tit"
>Endere&ccedil;o Residencial</td>
                        </tr>
                        <tr class="tb_corpo"> 
                          <td height="10"> <table width="100%" border="0" cellspacing="0">
        <tr> 
          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
          <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
          <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
 <%if isnull(rua_res) or rua_res="" then%>
             <input name="rua_res_fam" type="hidden" class="borda" id="rua_res_fam" value="" size="30">
 <%else
           response.write(Server.URLEncode(rua_res))%>
            <input name="rua_res_fam" type="hidden" class="borda" id="rua_res_fam" value="<%response.write(Server.URLEncode(rua_res))%>" size="30">
<%end if%>			
            </font></td>
          <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
          <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
          <td width="206" class="tb_corpo"
><font class="form_corpo"> 
<%response.write(num_res)%>
<input name="num_res_fam" type="hidden" class="borda" id="num_res_fam"  value="<%response.write(num_res)%>" size="12" maxlength="10">
            </font></td>
          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
          <td width="15" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
          <td width="139" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> 
 <%if isnull(comp_res) or comp_res="" then%>
              <input name="comp_res_fam" type="hidden" class="borda" id="comp_res" value="" size="12" maxlength="10">
 <%else
response.write(Server.URLEncode(comp_res))%>
              <input name="comp_res_fam" type="hidden" class="borda" id="comp_res" value="<%response.write(Server.URLEncode(comp_res))%>" size="12" maxlength="10">
<%end if%>


              </font></div></td>
        </tr>
        <tr> 
          <td width="145" height="21" class="tb_corpo"
><font class="form_dado_texto">Estado</font></td>
          <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
          <td width="217" height="21" class="tb_corpo"
><font class="form_corpo">
              <%
if isnull(estadores)or estadores="" then
%>
              <input name="estadores_fam" type="hidden" class="borda" id="estadores_fam" value="RJ" size="12" maxlength="10">
<%
else			  				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF where SG_UF='"&estadores&"'"
		RS2.Open SQL2, CON0

NO_UF= RS2("NO_UF")

response.Write(Server.URLEncode(NO_UF))
%>
              <input name="estadores_fam" type="hidden" class="borda" id="estadores_fam" value="<%response.write(Server.URLEncode(estadores))%>" size="12" maxlength="10">
<%
END IF
%>


            </font></td>
          <td width="140" height="21" class="tb_corpo"
><font class="form_dado_texto">Cidade</font></td>
          <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
          <td width="206" class="tb_corpo"
><font class="form_corpo"> 
                <%
if isnull(estadores)or estadores="" or cidres="" or isnull(cidres) then
%>
              <input name="cidres_fam" type="hidden" class="borda" id="cidres_fam" value="6001" size="12" maxlength="10">
<%
else	
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='"&estadores&"' And CO_Municipio="&cidres
		RS2m.Open SQL2m, CON0
		
NO_UF= RS2m("NO_Municipio")
response.Write(Server.URLEncode(NO_UF))
%>
              <input name="cidres_fam" type="hidden" class="borda" id="cidres_fam" value="<%response.write(Server.URLEncode(cidres))%>" size="12" maxlength="10">
<%
END IF
%>

              </font></td>
          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Bairro</font></td>
          <td width="15" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
          <td width="139" height="21" class="tb_corpo"
><font class="form_corpo"> 
                <%
if isnull(estadores)or estadores="" or cidres="" or isnull(cidres) or bairrores="" or isnull(bairrores)then
%>
              <input name="bairrores_fam" type="hidden" class="borda" id="bairrores_fam" value="100" size="12" maxlength="10">
<%
else	
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Bairro="&bairrores&" AND CO_Municipio="&cidres&" AND SG_UF='"&estadores&"' order by NO_Bairro"
		RS2b.Open SQL2b, CON0

if RS2b.eof then
%>		
	              <input name="bairrores_fam" type="hidden" class="borda" id="bairrores_fam" value="100" size="12" maxlength="10">	
<%else
CO_UF= RS2b("CO_Bairro")
NO_UF= RS2b("NO_Bairro")
response.Write(Server.URLEncode(NO_UF))
%>
              <input name="bairrores_fam" type="hidden" class="borda" id="bairrores_fam" value="<%response.write(Server.URLEncode(CO_UF))%>" size="12" maxlength="10">
<%
END IF
END IF
%>

 </font></td>
        </tr>
        <tr> 
          <td width="145" height="10" class="tb_corpo"
><font class="form_dado_texto">CEP</font></td>
          <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
          <td width="217" height="10" class="tb_corpo"
><font class="form_dado_texto"> <%response.write(cep)%>
            <input name="cep_fam" type="hidden" class="borda" id="cep_fam" value="<%response.write(cep)%>" size="11" maxlength="9">

            </font></td>
          <td width="140" height="10" class="tb_corpo"
>&nbsp;</td>
          <td width="19" class="tb_corpo"
>&nbsp;</td>
          <td width="206" class="tb_corpo"
>&nbsp;</td>
          <td width="90" class="tb_corpo"
>&nbsp;</td>
          <td width="15" class="tb_corpo"
>&nbsp; </td>
          <td width="139" height="10" class="tb_corpo"
>&nbsp;</td>
        </tr>
        <tr> 
          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">Telefones deste endere&ccedil;o</font></div></td>
          <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
          <td height="10" colspan="2" class="tb_corpo"
><font class="form_corpo"><%response.write(tel_res)%> 
            <input name="tel_res_fam" type="hidden" class="borda" id="tel_res_fam" value="<%response.write(tel_res)%>" size="50" maxlength="50">

            </font> </td>
          <td width="19" class="tb_corpo"
> <div align="center"></div></td>
          <td width="206" class="tb_corpo"
>&nbsp;</td>
          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Mesmo endere&ccedil;o do aluno</font></td>
          <td width="15" class="tb_corpo"
><div align="center"><font class="form_dado_texto">:</font> </div></td>
          <td width="139" height="10" class="tb_corpo"
>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr> 
                                      <td width="9%"><input type="radio" name="mes_end" value="s"  onClick="recuperarEnd('<%response.Write(cod)%>','<%response.Write(foco)%>');BD_aux(this.value,cod_consulta.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'ID_Res_Aluno')" checked></td>
                                      <td width="25%"><font class="form_corpo">Sim</font></td>
                                      <td width="5%"><input name="mes_end" type="radio"  onClick="recuperarOrigemEnd('<%response.Write(cod)%>','<%response.Write(foco)%>');BD_aux(this.value,cod_consulta.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'ID_Res_Aluno')" value="n"></td>
                                      <td width="61%"><font class="form_corpo">N&atilde;o 
                                        
                  <input name="id_res_fam_aux" type="hidden" id="id_res_fam_aux" value="s">
                                        </font></td>
                                    </tr>
                                  </table>
</td>
        </tr>
      </table></td>
                        </tr>
                      </table>
<%
'origem endereços
elseif opt="oe" then
cod= request.form("cod_pub")
foco= request.form("foco_pub")
session("id_res_familiar")="n"
%>

<table width="100%" border="0" cellspacing="0" cellpadding="0">
                        <tr class="tb_corpo"> 
                          <td class="tb_tit"
>Endere&ccedil;o Residencial</td>
                        </tr>
                        <tr class="tb_corpo"> 
                          <td height="10"> <table width="100%" border="0" cellspacing="0">
        <tr> 
          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
          <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
          <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
            <input name="rua_res_fam" type="text" class="borda" id="rua_res" size="30" maxlength="60">
            </font></td>
          <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
          <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
          <td width="206" class="tb_corpo"
><font class="form_corpo"> 
            <input name="num_res_fam" type="text" class="borda" id="num_res_fam" size="12" maxlength="10" onBlur="ValidaNumResFam(this.value)">
            </font></td>
          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
          <td width="15" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
          <td width="139" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
              <input name="comp_res_fam" type="text" class="borda" id="comp_res" size="12" maxlength="30">
              </font></div></td>
        </tr>
        <tr> 
          <td width="145" height="21" class="tb_corpo"
><font class="form_dado_texto">Estado</font></td>
          <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
          <td width="217" height="21" class="tb_corpo"
><font class="form_corpo"> 

              <font class="form_corpo">
              <select name="select" class="borda" id="select" onChange="recuperarCidResFam(this.value)">
                <option value="0" selected> </option>
                <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")

%>
                <option value="<%=SG_UF%>"> 
                <% response.Write(Server.URLEncode(NO_UF))%>
                </option>
                <%	
RS2.MOVENEXT
WEND
%>
              </select></font> </font></td>
          <td width="140" height="21" class="tb_corpo"
><font class="form_dado_texto">Cidade</font></td>
          <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
          <td width="206" class="tb_corpo"
><font class="form_corpo"> 
            <div id="cid_res_fam"> 
              <select name="cidres_fam" class="borda" id="select10" onChange="recuperarBairroResFam(this.value)">
                <option value="0" selected> </option>
              </select>
            </div>
            </font></td>
          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Bairro</font></td>
          <td width="15" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
          <td width="139" height="21" class="tb_corpo"
><font class="form_corpo">            <div id="bairro_res_fam"> 
              <select name="bairrores_fam" class="borda" id="bairrores">
                <option value="0" selected> </option>
              </select></div> </font></td>
        </tr>
        <tr> 
          <td width="145" height="10" class="tb_corpo"
><font class="form_dado_texto">CEP</font></td>
          <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
          <td width="217" height="10" class="tb_corpo"
><font class="form_dado_texto"> 
            <input name="cep_fam" type="text" class="borda" id="cep" onKeyup="formatar(this, '#####-###')" size="11" maxlength="9" onBlur="ValidaCepResFam(this.value)">
            </font></td>
          <td width="140" height="10" class="tb_corpo"
>&nbsp;</td>
          <td width="19" class="tb_corpo"
>&nbsp;</td>
          <td width="206" class="tb_corpo"
>&nbsp;</td>
          <td width="90" class="tb_corpo"
>&nbsp;</td>
          <td width="15" class="tb_corpo"
>&nbsp; </td>
          <td width="139" height="10" class="tb_corpo"
>&nbsp;</td>
        </tr>
        <tr> 
          <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">Telefones deste endere&ccedil;o</font></div></td>
          <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
          <td height="10" colspan="2" class="tb_corpo"
><font class="form_corpo"> 
            <input name="tel_res_fam" type="text" class="borda" id="tel_res" size="50" maxlength="100">
            </font> </td>
          <td width="19" class="tb_corpo"
> <div align="center"></div></td>
          <td width="206" class="tb_corpo"
>&nbsp;</td>
          <td width="90" class="tb_corpo"
><font class="form_dado_texto">Mesmo endere&ccedil;o do aluno</font></td>
          <td width="15" class="tb_corpo"
><div align="center"><font class="form_dado_texto">:</font> </div></td>
          <td width="139" height="10" class="tb_corpo"
>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
                                    <tr> 
                                      <td width="9%"><input type="radio" name="mes_end" value="s"  onClick="recuperarEnd('<%response.Write(cod)%>','<%response.Write(foco)%>');BD_aux(this.value,cod_consulta.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'ID_Res_Aluno')"></td>
                                      <td width="25%"><font class="form_corpo">Sim</font></td>
                                      <td width="5%"><input name="mes_end" type="radio"  onClick="recuperarOrigemEnd('<%response.Write(cod)%>','<%response.Write(foco)%>');BD_aux(this.value,cod_consulta.value,cod_familiar.value,tp_vinc_familiar_aux.value,co_vinc_familiar_aux.value,'ID_Res_Aluno')" value="n" checked></td>
                                      <td width="61%"><font class="form_corpo">N&atilde;o 
                                        
                  <input name="id_res_fam_aux" type="hidden" id="id_res_fam_aux" value="n">
                                        </font></td>
                                    </tr>
                                  </table>
										</td>
        </tr>
      </table></td>
                        </tr>
                      </table>
<%end if%>