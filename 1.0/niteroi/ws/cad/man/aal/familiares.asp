 cellspacing="0">
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
                                            <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                                              <%response.write(num_res)%>
                                              <input name="num_res_fam" type="hidden" class="borda" id="num_res_fam"  value="<%response.write(num_res)%>" size="12" maxlength="10">
                                              </font></td>
                                            <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                                            <td width="11" class="tb_corpo"
> 
                                              <div align="center"><font class="form_dado_texto">:</font></div></td>
                                            <td width="149" height="10" class="tb_corpo"
> 
                                              <div align="left"><font class="form_corpo"> 
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
                                            <td width="196" class="tb_corpo"
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
                                            <td width="11" class="tb_corpo"
> 
                                              <div align="center"><font class="form_dado_texto">:</font></div></td>
                                            <td width="149" height="21" class="tb_corpo"
><font class="form_corpo"> 
                                              <%
if isnull(estadores)or estadores="" or cidres="" or isnull(cidres) or bairrores="" or isnull(bairrores)then
%>
                                              <input name="bairrores_fam" type="hidden" class="borda" id="bairrores_fam" value="6001" size="12" maxlength="10">
                                              <%
else	
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Bairro="&bairrores&" AND CO_Municipio="&cidres&" AND SG_UF='"&estadores&"' order by NO_Bairro"
		RS2b.Open SQL2b, CON0
CO_UF= RS2b("CO_Bairro")
NO_UF= RS2b("NO_Bairro")
response.Write(Server.URLEncode(NO_UF))
%>
                                              <input name="bairrores_fam" type="hidden" class="borda" id="bairrores_fam" value="<%response.write(Server.URLEncode(CO_UF))%>" size="12" maxlength="10">
                                              <%
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
><font class="form_dado_texto"> 
                                              <%response.write(cep)%>
                                              <input name="cep_fam" type="hidden" class="borda" id="cep_fam" value="<%response.write(cep)%>" size="11" maxlength="9">
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
>&nbsp; </td>
                                            <td width="149" height="10" class="tb_corpo"
>&nbsp;</td>
                                          </tr>
                                          <tr> 
                                            <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">Telefones deste endere&ccedil;o</font></div></td>
                                            <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                                            <td height="10" colspan="2" class="tb_corpo"
><font class="form_corpo"> 
                                              <%response.write(tel_res)%>
                                              <input name="tel_res_fam" type="hidden" class="borda" id="tel_res_fam" value="<%response.write(tel_res)%>" size="50" maxlength="50">
                                              </font> </td>
                                            <td width="19" class="tb_corpo"
> <div align="center"></div></td>
                                            <td width="196" class="tb_corpo"
>&nbsp;</td>
                                            <td width="90" class="tb_corpo"
><font class="form_dado_texto">Mesmo endere&ccedil;o do aluno</font></td>
                                            <td width="11" class="tb_corpo"
>
<div align="center"><font class="form_dado_texto">:</font> </div></td>
                                            <td width="149" height="10" class="tb_corpo"
> 
                                              <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                                <tr> 
                                                  <td width="9%"><input type="radio" name="mes_end" id="mes_end" value="s"  onClick="recuperarEnd('<%response.Write(cod)%>','<%response.Write(foco)%>')" checked></td>
                                                  <td width="25%"><font class="form_corpo">Sim</font></td>
                                                  <td width="5%"><input name="mes_end" id="mes_end"  type="radio"  onClick="recuperarOrigemEnd('<%response.Write(cod)%>','<%response.Write(foco)%>')" value="n"></td>
                                                  <td width="61%"><font class="form_corpo">N&atilde;o 
                                                    <input name="id_res_fam_aux" type="hidden" id="id_res_fam_aux" value="s">
                                                    </font></td>
                                                </tr>
                                              </table></td>
                                          </tr>
                                        </table></td>
                                    </tr>
                                  </table>
                                  <%else%>
                                  <table width="100%" border="0" cellspacing="0">
                                    <tr> 
                                      <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                                      <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                                      <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                                        <%if rua_res="" or isnull(rua_res) then%>
                                        <input name="rua_res_fam" type="text" class="borda" id="rua_res_fam" size="30" maxlength="60">
                                        <%else%>
                                        <input name="rua_res_fam" type="text" class="borda" id="rua_res_fam" value="<%response.write(Server.URLEncode(rua_res))%>" size="30" maxlength="60">
                                        <%end if%>
                                        </font></td>
                                      <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                                      <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                                      <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                                        <input name="num_res_fam" type="text" class="borda" id="num_res_fam"  value="<%response.write(num_res)%>" size="12" maxlength="10" onBlur="ValidaNumResFam(this.value)">
                                        </font></td>
                                      <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                                      <td width="11" class="tb_corpo"
> 
                                        <div align="center"><font class="form_dado_texto">:</font></div></td>
                                      <td width="149" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
                                          <%if comp_res="" or isnull(comp_res) then%>
                                          <input name="comp_res_fam" type="text" class="borda" id="comp_res_fam" size="20" maxlength="30">
                                          <%else%>
                                          <input name="comp_res_fam" type="text" class="borda" id="comp_res_fam" value="<%response.write(Server.URLEncode(comp_res))%>" size="20" maxlength="30">
                                          <%end if%>
                                          </font></div></td>
                                    </tr>
                                    <tr> 
                                      <td width="145" height="21" class="tb_corpo"
><font class="form_dado_texto">Estado</font></td>
                                      <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                                      <td width="217" height="21" class="tb_corpo"
><font class="form_corpo"> <font class="form_corpo"> 
                                        <select name="estadores_fam" class="borda" id="estadores_fam" onChange="recuperarCidResFam(this.value)">
                                          <option value="0" > </option>
                                          <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")

if SG_UF = estadores then
%>
                                          <option value="<%=SG_UF%>" selected> 
                                          <% response.Write(Server.URLEncode(NO_UF))%>
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
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                                      <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                                        <div id="cid_res_fam"> 
                                          <select name="cidres_fam" class="borda" id="cidres_fam" onChange="recuperarBairroResFam(estadores_fam.value,this.value)">
                                            <%
if isnull(estadores) or estadores="" then
%>
                                            <option value="0" selected> </option>
                                            <%else%>
                                            <option value="0" > </option>
                                            <%																		  
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='"&estadores&"' order by NO_Municipio"
		RS2m.Open SQL2m, CON0
		
while not RS2m.EOF						
SG_UF= RS2m("CO_Municipio")
NO_UF= RS2m("NO_Municipio")

if SG_UF = cidres then
%>
                                            <option value="<%=SG_UF%>" selected> 
                                            <% response.Write(Server.URLEncode(NO_UF))%>
                                            </option>
                                            <% else %>
                                            <option value="<%=SG_UF%>"> 
                                            <% response.Write(Server.URLEncode(NO_UF))%>
                                            </option>
                                            <%
end if	
RS2m.MOVENEXT
WEND
end if
%>
                                          </select>
                                        </div>
                                        </font></td>
                                      <td width="90" class="tb_corpo"
><font class="form_dado_texto">Bairro</font></td>
                                      <td width="11" class="tb_corpo"
> 
                                        <div align="center"><font class="form_dado_texto">:</font></div></td>
                                      <td width="149" height="21" class="tb_corpo"
> <div id="bairro_res_fam"><font class="form_corpo"> 
                                          <select name="bairrores_fam" class="borda" id="bairrores_fam">
                                            <%
if isnull(estadores) or estadores="" or isnull(cidres) or cidres="" then
%>
                                            <option value="0"> </option>
                                            <%else
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Municipio="&cidres&" AND SG_UF='"&estadores&"' order by NO_Bairro"
		RS2b.Open SQL2b, CON0

IF RS2b.EOF then
%>
                                            <option value="0"> 
                                            <% response.Write(Server.URLEncode("Bairros não cadastrados"))%>
                                            </option>
                                            <%else	
		
while not RS2b.EOF						
SG_UF= RS2b("CO_Bairro")
NO_UF= RS2b("NO_Bairro")
if SG_UF = bairrores then
%>
                                            <option value="<%=SG_UF%>" selected> 
                                            <% response.Write(Server.URLEncode(NO_UF))%>
                                            </option>
                                            <% else %>
                                            <option value="<%=SG_UF%>"> 
                                            <% response.Write(Server.URLEncode(NO_UF))%>
                                            </option>
                                            <%
end if	

RS2b.MOVENEXT
WEND
end if
end if
%>
                                          </select>
                                          </font></div></td>
                                    </tr>
                                    <tr> 
                                      <td width="145" height="10" class="tb_corpo"
><font class="form_dado_texto">CEP</font></td>
                                      <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                                      <td width="217" height="10" class="tb_corpo"
><font class="form_dado_texto"> 
                                        <input name="cep_fam" type="text" class="borda" id="cep_fam" onKeyup="formatar(this, '#####-###')" value="<%response.write(cep)%>" size="11" maxlength="9" onBlur="ValidaCepResFam(this.value)">
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
>&nbsp; </td>
                                      <td width="149" height="10" class="tb_corpo"
>&nbsp;</td>
                                    </tr>
                                    <tr> 
                                      <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">Telefones deste endere&ccedil;o</font></div></td>
                                      <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                                      <td height="10" colspan="2" class="tb_corpo"
><font class="form_corpo"> 
                                        <input name="tel_res_fam" type="text" class="borda" id="tel_res_fam" value="<%response.write(tel_res)%>" size="42" maxlength="100">
                                        </font> </td>
                                      <td width="19" class="tb_corpo"
> <div align="center"></div></td>
                                      <td width="196" class="tb_corpo"
> </td>
                                      <td width="90" class="tb_corpo"
><font class="form_dado_texto">Mesmo endere&ccedil;o do aluno</font></td>
                                      <td width="11" class="tb_corpo"
><div align="center"><font class="form_dado_texto">:</font> </div></td>
                                      <td width="149" height="10" class="tb_corpo"
> <table width="100%" border="0" cellspacing="0" cellpadding="0">
                                          <tr> 
                                            <td width="9%"><input type="radio" name="mes_end" value="s"  onClick="recuperarEnd('<%response.Write(cod_consulta)%>','<%response.Write(foco)%>')"></td>
                                            <td width="25%"><font class="form_corpo">Sim</font></td>
                                            <td width="5%"><input name="mes_end" type="radio"  onClick="recuperarOrigemEnd('<%response.Write(cod_consulta)%>','<%response.Write(foco)%>')" value="n" checked></td>
                                            <td width="61%"><font class="form_corpo">N&atilde;o 
                                              <input name="id_res_fam_aux" type="hidden" id="id_res_fam_aux" value="n">
                                              </font></td>
                                          </tr>
                                        </table></td>
                                    </tr>
                                  </table>
                                  <%end if%>
                                </td>
                              </tr>
                            </table>
                          </div> 
                          <table width="100%" border="0" cellspacing="0" dwcopytype="CopyTableRow">
                            <tr> 
                              <td height="10" colspan="9" class="tb_tit"
><div align="left">Endere&ccedil;o Comercial </div></td>
                            </tr>
                            <tr> 
                              <td width="145" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> Logradouro</font></div></td>
                              <td width="13" class="tb_corpo"
> <div align="left"><font class="form_dado_texto">:</font></div></td>
                              <td width="217" height="10" class="tb_corpo"
><font class="form_corpo"> 
                                <%if rua_com="" or isnull(rua_com) then%>
                                <input name="rua_com_fam" type="text" class="borda" id="rua_com_fam" size="30" maxlength="60">
                                <%else%>
                                <input name="rua_com_fam" type="text" class="borda" id="rua_com_fam" value="<%response.write(Server.URLEncode(rua_com))%>" size="30" maxlength="60">
                                <%end if%>
                                </font></td>
                              <td width="140" height="10" class="tb_corpo"
> <div align="left"><font class="form_dado_texto"> N&uacute;mero</font></div></td>
                              <td width="19" class="tb_corpo"
> <div align="center"><font class="form_dado_texto">:</font></div></td>
                              <td width="196" class="tb_corpo"
><font class="form_corpo"> 
                                <input name="num_com_fam" type="text" class="borda" id="num_com_fam" value="<%response.write(num_com)%>" size="12" maxlength="10" onBlur="ValidaNumComFam(this.value)">
                                </font></td>
                              <td width="90" class="tb_corpo"
><font class="form_dado_texto">Complemento</font></td>
                              <td width="11" class="tb_corpo"
>
<div align="center"><font class="form_dado_texto">:</font></div></td>
                              <td width="149" height="10" class="tb_corpo"
> <div align="left"><font class="form_corpo"> </font> <font class="form_corpo"> 
                                  <% if isnull(comp_com) or comp_com="" then%>
                                  <input name="comp_com_fam" type="text" class="borda" id="comp_com" size="20" maxlength="30">
                                  <%else%>
                                  <input name="comp_com_fam" type="text" class="borda" id="comp_com" value="<%response.write(Server.URLEncode(comp_com))%>" size="20" maxlength="30">
                                  <%end if%>
                                  </font></div></td>
                            </tr>
                            <tr class="tb_corpo"
> 
                              <td width="145" height="26"><font class="form_dado_texto">Estado</font></td>
                              <td width="13"> <div align="left"><font class="form_dado_texto">:</font></div></td>
                              <td width="217" height="26"><font class="form_corpo"> 
                                <select name="estadocom_fam" class="borda" id="estadocom_fam" onChange="recuperarCidComFam(this.value)">
                                  <option value="0" > </option>
                                  <%				
		Set RS2 = Server.CreateObject("ADODB.Recordset")
		SQL2 = "SELECT * FROM TB_UF order by NO_UF"
		RS2.Open SQL2, CON0
		
while not RS2.EOF						
SG_UF= RS2("SG_UF")
NO_UF= RS2("NO_UF")
if isnull(uf_natural) then
uf_natural="RJ"
end if
if SG_UF = estadocom then
%>
                                  <option value="<%=SG_UF%>" selected> 
                                  <% response.Write(Server.URLEncode(NO_UF))%>
                                  </option>
                                  <%else%>
                                  <option value="<%=SG_UF%>"> 
                                  <% response.Write(Server.URLEncode(NO_UF))%>
                                  </option>
                                  <%end if						
RS2.MOVENEXT
WEND
%>
                                </select>
                                </font></td>
                              <td width="140" height="26"><font class="form_dado_texto">Cidade</font></td>
                              <td width="19"> <div align="center"><font class="form_dado_texto">:</font></div></td>
                              <td width="196"> <div id="cid_com_fam"> 
                                  <select name="cidcom_fam" class="borda" id="cidcom_fam" onChange="recuperarBairroComFam(estadocom_fam.value,this.value)">
                                    <%
if isnull(estadocom) or estadocom="" then
%>
                                    <option value="0" selected> </option>
                                    <%else %>
                                    <option value="0" > </option>
                                    <%																  
Set RS2m = Server.CreateObject("ADODB.Recordset")
		SQL2m = "SELECT * FROM TB_Municipios WHERE SG_UF='"&estadocom&"' order by NO_Municipio"
		RS2m.Open SQL2m, CON0
		
while not RS2m.EOF						
SG_UF= RS2m("CO_Municipio")
NO_UF= RS2m("NO_Municipio")

if SG_UF = cidcom then
%>
                                    <option value="<%=SG_UF%>" selected> 
                                    <% response.Write(Server.URLEncode(NO_UF))%>
                                    </option>
                                    <% else %>
                                    <option value="<%=SG_UF%>"> 
                                    <% response.Write(Server.URLEncode(NO_UF))%>
                                    </option>
                                    <%
end if	
RS2m.MOVENEXT
WEND
end if
%>
                                  </select>
                                </div></td>
                              <td width="90"><font class="form_dado_texto">Bairro</font></td>
                              <td width="11">
<div align="center"><font class="form_dado_texto">:</font></div></td>
                              <td width="149" height="26"> <div id="bairro_com_fam"><font class="form_corpo"> 
                                  <select name="bairrocom_fam" class="borda" id="bairrocom_fam">
                                    <%
if isnull(estadocom) or estadocom="" or isnull(cidcom) or cidcom="" then
%>
                                    <option value="0" selected> </option>
                                    <%else
Set RS2b = Server.CreateObject("ADODB.Recordset")
		SQL2b = "SELECT * FROM TB_Bairros WHERE CO_Municipio="&cidcom&" AND SG_UF='"&estadocom&"' order by NO_Bairro"
		RS2b.Open SQL2b, CON0

IF RS2b.EOF then
%>
                                    <option value="0"> 
                                    <% response.Write(Server.URLEncode("Bairros não cadastrados"))%>
                                    </option>
                                    <%else	
while not RS2b.EOF						
SG_UF= RS2b("CO_Bairro")
NO_UF= RS2b("NO_Bairro")
if SG_UF = bairrocom then
%>
                                    <option value="<%=SG_UF%>" selected> 
                                    <% response.Write(Server.URLEncode(NO_UF))%>
                                    </option>
                                    <% else %>
                                    <option value="<%=SG_UF%>"> 
                                    <% response.Write(Server.URLEncode(NO_UF))%>
                                    </option>
                                    <%
end if
RS2b.MOVENEXT
WEND
end if
end if
%>
                                  </select>
                                  </font> </div></td>
                            </tr>
                            <tr class="tb_corpo"
> 
                              <td width="145" height="26"><font class="form_dado_texto">CEP</font></td>
                              <td width="13"> <div align="left"><font class="form_dado_texto">:</font></div></td>
                              <td width="217" height="26"><font class="form_dado_texto"> 
                                <input name="cepcom_fam" type="text" class="borda" id="cepcom_fam" onKeyup="formatar(this, '#####-###')" value="<%response.write(cepcom)%>" size="11" maxlength="9" onBlur="ValidaCepComFam(this.value)">
                                </font></td>
                              <td width="140" height="26">&nbsp;</td>
                              <td width="19">&nbsp;</td>
                              <td width="196">&nbsp;</td>
                              <td width="90">&nbsp;</td>
                              <td width="11">&nbsp;</td>
                              <td width="149" height="26">&nbsp;</td>
                            </tr>
                            <tr class="tb_corpo"
> 
                              <td width="145" height="28"> <div align="left"><font class="form_dado_texto">Telefones 
                                  deste endere&ccedil;o<font class="form_dado_texto">:</font></font></div></td>
                              <td width="13"> <div align="left"><font class="form_dado_texto">:</font></div></td>
                              <td height="28" colspan="2"><font class="form_corpo"> 
                                <input name="tel_com_fam" type="text" class="borda" id="tel_com_fam" value="<%response.write(tel_com)%>" size="42" maxlength="100">
                                </font> </td>
                              <td width="19">&nbsp; </td>
                              <td width="196">&nbsp;</td>
                              <td width="90">&nbsp;</td>
                              <td width="11">&nbsp;</td>
                              <td width="149" height="28">&nbsp;</td>
                            </tr>
                          </table></td>
                      </tr>
                    </table>
<%end if%>
</div></td>
  </tr>
</table>
</td>
                  </tr>
                </table>
</td>
                  </tr>
<%
'end if do if abre_campos
 'end if%>				  
<div id="responsaveis">
<% if aluno_vinculado="s" then%>
    <tr> 
      <td valign="bottom"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="10" colspan="9" class="tb_tit">Respons&aacute;veis</td>
          </tr>
          <tr> 
            <td width="145" height="10"><font class="form_dado_texto">Financeiro</font></td>
            <td width="13" height="10"> <div align="left"><font class="form_dado_texto">:</font></div></td>
            <td width="217" height="10"> <div align="left"><font class="form_dado_texto"> 
                <%

		Set RSRESPs = Server.CreateObject("ADODB.Recordset")
		SQLs = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="&codigo_aluno_vinculado
		RSRESPs.Open SQLs, CON1_AUX
		

if RSRESPs.EOF then
else
			resp_fin= RSRESPs("TP_Resp_Fin")

			co_vinc_familiar_aux=cod_consulta

		le ="n"
		while le ="n" 

		'response.Write("<br>SELECT * FROM TBI_Contatos WHERE TP_Contato='"&resp_fin&"' and CO_Matricula ="&co_vinc_familiar_aux)
		Set RSRESP_PED_vinc = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED_vinc= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&resp_fin&"' and CO_Matricula ="&codigo_aluno_vinculado
		RSRESP_PED_vinc.Open SQLRESP_PED_vinc, CONCONT_aux
			
			co_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("CO_Matricula_Vinc")
			tp_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("TP_Contato_Vinc")
			id_familia=RSRESP_PED_vinc("ID_Familia")
			id_end_bloq=RSRESP_PED_vinc("ID_End_Bloqueto")			
	
			if (isnull(tp_vinc_familiar_aux_mais_um) or tp_vinc_familiar_aux_mais_um="") and (isnull(co_vinc_familiar_aux_mais_um) or co_vinc_familiar_aux_mais_um="") then
			le="s"
			nome_familiar=RSRESP_PED_vinc("NO_Contato")
			'response.Write(" - '"&nome_familiar&"'")
			else
			co_vinc_familiar_aux=co_vinc_familiar_aux_mais_um
			resp_fin=tp_vinc_familiar_aux_mais_um
			end if
		wend 	
end if				  
		if nome_familiar="" or isnull(nome_familiar) then
		nome_familiar="Familiar "&resp_fin&" sem nome cadastrado"
		end if


response.Write(Server.URLEncode(nome_familiar))
%>
                <input name="rf" type="hidden" id="rf" value="<%response.Write(resp_fin)%>">
                </font></div></td>
            <td width="140" height="10"><font class="form_dado_texto">Fam&iacute;lia</font></td>
            <td width="19" height="10"><div align="center"><font class="form_dado_texto">:</font></div></td>
            <td width="196" height="10"><div align="left"><font class="form_dado_texto"> 
                <%response.Write(id_familia)%>
                <input name="id_familia" type="hidden" id="id_familia" value="<%response.Write(id_familia)%>">
                </font></div></td>
            <td width="90"><font class="form_dado_texto">End. Bloqueto </font></td>
            <td width="11"><div align="center"><font class="form_dado_texto">?</font></div></td>
            <td width="149"><div align="left"><font class="form_dado_texto"> 
                <%if id_end_bloq="R" then
						  %>
                Residencial 
                <%
elseif id_end_bloq="C" then%>
                <input name="bloq" type="hidden" id="bloq" value="<%response.Write(id_end_bloq)%>">
                Comercial 
                <%else
end if
%>
                </font></div></td>
          </tr>
          <tr> 
            <td width="145" height="10"><font class="form_dado_texto">Pedag&oacute;gico</font></td>
            <td width="13" height="10"> <div align="left"><font class="form_dado_texto">:</font></div></td>
            <td width="217" height="10"> 
              <div align="left"><font class="form_dado_texto"> 
                <%

		Set RSRESPs = Server.CreateObject("ADODB.Recordset")
		SQLs = "SELECT * FROM TBI_Alunos WHERE CO_Matricula ="&codigo_aluno_vinculado
		RSRESPs.Open SQLs, CON1_AUX
		

if RSRESPs.EOF then
else
			resp_ped= RSRESPs("TP_Resp_Ped")
			tp_vinc_familiar_aux_mais_um=resp_fin

		le ="n"
		while le ="n" 

		'response.Write("<br>SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc)
		Set RSRESP_PED_vinc = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED_vinc= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&resp_ped&"' and CO_Matricula ="&codigo_aluno_vinculado
		RSRESP_PED_vinc.Open SQLRESP_PED_vinc, CONCONT_aux
			
			co_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("CO_Matricula_Vinc")
			tp_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("TP_Contato_Vinc")
			id_end_circ=RSRESP_PED_vinc("ID_End_Bloqueto")
	
			if (isnull(tp_vinc_familiar_aux_mais_um) or tp_vinc_familiar_aux_mais_um="") and (isnull(co_vinc_familiar_aux_mais_um) or co_vinc_familiar_aux_mais_um="") then
			le="s"
			nome_familiar=RSRESP_PED_vinc("NO_Contato")
			'response.Write(" - '"&nome_familiar&"'")
			else
			co_vinc_familiar_aux=co_vinc_familiar_aux_mais_um
			resp_ped=tp_vinc_familiar_aux_mais_um
			end if
		wend 	
end if				  
		if nome_familiar="" or isnull(nome_familiar) then
		nome_familiar="Familiar "&resp_ped&" sem nome cadastrado"
		end if

response.Write(Server.URLEncode(nome_familiar))
%>
                <input name="rp" type="hidden" id="rp" value="<%response.Write(resp_ped)%>">
                </font></div></td>
            <td width="140" height="10">&nbsp;</td>
            <td width="19" height="10">&nbsp;</td>
            <td width="196" height="10">&nbsp;</td>
            <td width="90" height="10"><font class="form_dado_texto">End. Circular</font></td>
            <td width="11" height="10"><div align="center"><font class="form_dado_texto">?</font></div></td>
            <td width="149" height="10"><div align="left"><font class="form_dado_texto"> 
                <%if id_end_circ="R" then%>
                Residencial 
                <%
elseif id_end_circ="C" then%>
                <input name="circ" type="hidden" id="circ" value="<%response.Write(id_end_circ)%>">
                Comercial 
                <%else
end if
%>
                </font></div></td>
          </tr>
        </table>
</td>
</tr>  
<% else %>
    <tr> 
      <td valign="bottom"> 
        <table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr> 
            <td height="10" colspan="9" class="tb_tit">Respons&aacute;veis</td>
          </tr>
          <tr> 
            <td width="145" height="10"><font class="form_dado_texto">Financeiro</font></td>
            <td width="13" height="10"> <div align="left"><font class="form_dado_texto">:</font></div></td>
            <td width="217" height="10"> <select name="rf" class="borda" onChange="GravaResponsaveis(this.value,'TP_Resp_Fin',0,'TP_Resp_Fin','<%response.write(cod_consulta)%>')">
                <option value="0" ></option>
                <%

		Set RSCONTPR = Server.CreateObject("ADODB.Recordset")
		SQLCONTPR = "SELECT * FROM TB_Tipo_Contatos Order by NU_Prioridade_Combo"
		RSCONTPR.Open SQLCONTPR, CON0
total_tp_familiares=0		
while not RSCONTPR.EOF	  
cod_familiar = RSCONTPR("TP_Contato")

		Set RSRESP_PED = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod
		RSRESP_PED.Open SQLRESP_PED, CONCONT_aux
		
if RSRESP_PED.EOF then
else
cod_vinc=RSRESP_PED("CO_Matricula_Vinc")
tp_familiar_vinc=RSRESP_PED("TP_Contato_Vinc")
nome_familiar=RSRESP_PED("NO_Contato")



if (isnull(cod_vinc) or cod_vinc="NULL" or cod_vinc="") and (isnull(tp_familiar_vinc) or tp_familiar_vinc="NULL" or tp_familiar_vinc="") then
else
		le ="n"
		while le ="n" 

		'response.Write("<br>SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc)
		Set RSRESP_PED_vinc = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED_vinc= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc
		RSRESP_PED_vinc.Open SQLRESP_PED_vinc, CONCONT_aux
			
			co_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("CO_Matricula_Vinc")
			tp_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("TP_Contato_Vinc")
	
			if (isnull(tp_vinc_familiar_aux_mais_um) or tp_vinc_familiar_aux_mais_um="") and (isnull(co_vinc_familiar_aux_mais_um) or co_vinc_familiar_aux_mais_um="") then
			le="s"
			nome_familiar=RSRESP_PED_vinc("NO_Contato")
			response.Write(" - '"&nome_familiar&"'")
			else
			cod_vinc=co_vinc_familiar_aux_mais_um
			tp_familiar_vinc=tp_vinc_familiar_aux_mais_um
			end if
		wend 	
end if				  
		if nome_familiar="" or isnull(nome_familiar) then
		nome_familiar="Familiar "&tp_familiar_vinc&" sem nome cadastrado"
		end if
		
if cod_familiar=resp_fin then
id_familia=RSRESP_PED("ID_Familia")
id_end_bloq=RSRESP_PED("ID_End_Bloqueto")

						  %>
                <option value="<%response.Write(cod_familiar)%>" selected> 
                <%response.Write(Server.URLEncode(nome_familiar))%>
                </option>
                <%
else
						  %>
                <option value="<%response.Write(cod_familiar)%>" > 
                <%response.Write(Server.URLEncode(nome_familiar))%>
                </option>
                <%
end if
end if

RSCONTPR.MOVENEXT
WEND	  
%>
              </select> </td>
            <td width="140" height="10"><font class="form_dado_texto">Fam&iacute;lia</font></td>
            <td width="19" height="10"><div align="center"><font class="form_dado_texto">:</font></div></td>
            <td width="196" height="10"><input name="id_familia" type="text" class="borda" id="rg2" onBlur="GravaResponsaveis(this.value,'ID_Familia',rf.value,'TP_Resp_Fin','<%response.write(cod_consulta)%>')" value="<%response.Write(id_familia)%>" size="30" maxlength="50"> 
            </td>
            <td width="90"><font class="form_dado_texto">End. Bloqueto </font></td>
            <td width="11"><div align="center"><font class="form_dado_texto">?</font></div></td>
            <td width="149"><select name="bloq" class="borda" id="bloq" onChange="GravaResponsaveis(this.value,'ID_End_Bloqueto',rf.value,'TP_Resp_Fin','<%response.write(cod_consulta)%>')">
                <%if id_end_bloq="R" then
						  %>
                <option value="R" selected> Residencial </option>
                <option value="C"> Comercial </option>
                <%
elseif id_end_bloq="C" then
						  %>
                <option value="R"> Residencial </option>
                <option value="C" selected> Comercial </option>
                <%else%>
                <option value="0" selected></option>
                <option value="R"> Residencial </option>
                <option value="C" > Comercial </option>
                <%
end if
%>
              </select> </td>
          </tr>
          <tr> 
            <td width="145" height="10"><font class="form_dado_texto">Pedag&oacute;gico</font></td>
            <td width="13" height="10"> <div align="left"><font class="form_dado_texto">:</font></div></td>
            <td width="217" height="10"> 
              <select name="rp" class="borda" onChange="GravaResponsaveis(this.value,'TP_Resp_Ped',0,'TP_Resp_Ped','<%response.write(cod)%>')">
                <option value="0" ></option>
                <%

		Set RSCONTPR = Server.CreateObject("ADODB.Recordset")
		SQLCONTPR = "SELECT * FROM TB_Tipo_Contatos Order by NU_Prioridade_Combo"
		RSCONTPR.Open SQLCONTPR, CON0
total_tp_familiares=0		
while not RSCONTPR.EOF	  
cod_familiar = RSCONTPR("TP_Contato")

		Set RSRESP_PED = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&cod_familiar&"' and CO_Matricula ="&cod
		RSRESP_PED.Open SQLRESP_PED, CONCONT_aux
		
if RSRESP_PED.EOF then
else
cod_vinc=RSRESP_PED("CO_Matricula_Vinc")
tp_familiar_vinc=RSRESP_PED("TP_Contato_Vinc")
nome_familiar=RSRESP_PED("NO_Contato")
if (isnull(cod_vinc) or cod_vinc="NULL" or cod_vinc="") and (isnull(tp_familiar_vinc) or tp_familiar_vinc="NULL" or tp_familiar_vinc="") then
else
		le ="n"
		while le ="n" 

		Set RSRESP_PED_vinc = Server.CreateObject("ADODB.Recordset")
		SQLRESP_PED_vinc= "SELECT * FROM TBI_Contatos WHERE TP_Contato='"&tp_familiar_vinc&"' and CO_Matricula ="&cod_vinc
		RSRESP_PED_vinc.Open SQLRESP_PED_vinc, CONCONT_aux
			
			co_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("CO_Matricula_Vinc")
			tp_vinc_familiar_aux_mais_um=RSRESP_PED_vinc("TP_Contato_Vinc")
	
			if (isnull(tp_vinc_familiar_aux_mais_um) or tp_vinc_familiar_aux_mais_um="") and (isnull(co_vinc_familiar_aux_mais_um) or co_vinc_familiar_aux_mais_um="") then
			le="s"
			nome_familiar=RSRESP_PED_vinc("NO_Contato")
			else
			cod_vinc=co_vinc_familiar_aux_mais_um
			tp_familiar_vinc=tp_vinc_familiar_aux_mais_um
			end if
		wend 
end if
		  if nome_familiar="" or isnull(nome_familiar) then
		  nome_familiar="Familiar "&tp_familiar_vinc&" sem nome cadastrado"
		  end if
						  
if cod_familiar=resp_ped then
id_familia=RSRESP_PED("ID_Familia")
id_end_bloq=RSRESP_PED("ID_End_Bloqueto")

						  %>
                <option value="<%response.Write(cod_familiar)%>" selected> 
                <%response.Write(Server.URLEncode(nome_familiar))%>
                </option>
                <%
else
						  %>
                <option value="<%response.Write(cod_familiar)%>" > 
                <%response.Write(Server.URLEncode(nome_familiar))%>
                </option>
                <%
end if
end if

RSCONTPR.MOVENEXT
WEND	  
%>
              </select> </td>
            <td width="140" height="10">&nbsp;</td>
            <td width="19" height="10">&nbsp;</td>
            <td width="196" height="10">&nbsp;</td>
            <td width="90" height="10"><font class="form_dado_texto">End. Circular</font></td>
            <td width="11" height="10"><div align="center"><font class="form_dado_texto">?</font></div></td>
            <td width="149" height="10"><select name="circ" class="borda" id="circ" onChange="GravaResponsaveis(this.value,'ID_End_Bloqueto',rp.value,'TP_Resp_Ped','<%response.write(cod)%>')">
                <%if id_end_bloq="R" then
						  %>
                <option value="R" selected> Residencial </option>
                <option value="C"> Comercial </option>
                <%
elseif id_end_bloq="C" then
						  %>
                <option value="R"> Residencial </option>
                <option value="C" selected> Comercial </option>
                <%else%>
                <option value="0" selected></option>
                <option value="R"> Residencial </option>
                <option value="C"> Comercial </option>
                <%
end if
%>
              </select></td>
          </tr>
        </table>
</td>
</tr> 
 <%end if%>
</div>
</table>				