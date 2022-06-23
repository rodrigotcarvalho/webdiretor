<!--#include file="../../../../inc/caminhos.asp"-->
<html xmlns="http://www.w3.org/1999/xhtml">
<head>
<meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
<title>Histórico do item</title>
<link href="../../../../estilos.css" rel="stylesheet" type="text/css">
</head>
<%
cod_item = request.QueryString("cod_item")

	Set CON9 = Server.CreateObject("ADODB.Connection") 
	ABRIR9 = "DBQ="& CAMINHO_ax & ";Driver={Microsoft Access Driver (*.mdb)}"
	CON9.Open ABRIR9	

	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT  * FROM TB_Item WHERE CO_Item ="& cod_item
	RS.Open SQL, CON9	

	nome  = RS("NO_Item")


%>
<body>
<table width="610" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td bgcolor="#FFFFFF"><table width="610" border="0" cellspacing="0" cellpadding="0">
      <tr class="form_dado_texto">
        <td colspan="4" align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr class="form_dado_texto">
            <td colspan="4" class="tb_tit">Histórico dos Lançamentos Efetuados </td>
            </tr>
          <tr class="form_dado_texto">
            <td class="form_corpo">&nbsp;</td>
            <td>&nbsp;</td>
            <td class="form_corpo">&nbsp;</td>
            <td>&nbsp;</td>
          </tr>
          <tr class="form_dado_texto">
            <td width="9%" class="form_corpo">Item:</td>
            <td width="11%"><%response.Write(cod_item)%></td>
            <td width="8%" class="form_corpo">Nome:</td>
            <td width="72%"><%response.Write(nome)%></td>
          </tr>
        </table></td>
      </tr>
      <tr class="form_dado_texto">
        <td align="center">&nbsp;</td>
        <td align="center">&nbsp;</td>
        <td align="center">&nbsp;</td>
        <td align="center">&nbsp;</td>
      </tr>
      <tr class="tb_tit">
        <td width="152" align="center">Tipo</td>
        <td width="152" align="center">Data da Operação</td>
        <td width="152" align="center">Quantidade</td>
        <td width="154" align="center">No de Controle</td>
      </tr>
      <%
  
	Set RSM = Server.CreateObject("ADODB.Recordset")
	SQLM = "SELECT 'Entrada' as Tipo, TB_NFiscais_Compra_Item.NU_NotaF as Pedido, TB_NFiscais_Compra_Item.QT_Item as QTD, TB_NFiscais_Compra.DA_NotaF as Data_BD FROM TB_NFiscais_Compra_Item, TB_NFiscais_Compra WHERE TB_NFiscais_Compra.NU_NotaF = TB_NFiscais_Compra_Item.NU_NotaF and CO_Item ="& cod_item &" UNION ALL SELECT 'Sa&iacute;da' as Tipo, TB_Mov_Estoque_Item.NU_Pedido as Pedido, TB_Mov_Estoque_Item.QT_Solicitado as QTD, TB_Mov_Estoque.DA_Pedido as Data_BD FROM TB_Mov_Estoque_Item,TB_Mov_Estoque WHERE TB_Mov_Estoque.NU_Pedido = TB_Mov_Estoque_Item.NU_Pedido and CO_Item ="& cod_item
	RSM.Open SQLM, CON9
	
	While not RSM.EOF
		tipo = RSM("Tipo")
	
		pedido = RSM("Pedido")
		quantidade = RSM("QTD")
		data_bd = RSM("Data_BD")
		
		data_split= Split(data_bd,"/")
		dia=data_split(0)
		mes=data_split(1)
		ano=data_split(2)
		
		
		dia=dia*1
		
		mes=mes*1
		hora=hora*1
		min=min*1
		
		if dia<10 then
		dia="0"&dia
		end if
		if mes<10 then
		mes="0"&mes
		end if
		da_show=dia&"/"&mes&"/"&ano
		
		  if tipo="Sa&iacute;da" then
			cor = "FF0000"
		  else
			cor = "0000FF"
		  end if		
		    
  %>
      <tr class="form_dado_texto">
        <td width="152" align="center"><%response.Write("<font color="&cor&">"&tipo&"</font>")%></td>
        <td width="152" align="center"><%response.Write(da_show)%></td>
        <td width="152" align="center"><%response.Write(quantidade)%></td>
        <td width="154" align="center"><%response.Write(pedido)%></td>
      </tr>
      <%
  RSM.MOVENEXT
  WEND
  
	Set RS = Server.CreateObject("ADODB.Recordset")
	SQL = "SELECT  * FROM TB_Item WHERE CO_Item ="& cod_item
	RS.Open SQL, CON9	

	estoque = RS("QT_Atual") 
 %>
      <tr class="tb_subtit">
        <td align="center">Saldo Atual</td>
        <td align="center">&nbsp;</td>
        <td align="center"><%response.Write(estoque)%></td>
        <td align="center">&nbsp;</td>
      </tr>
      <tr>
        <td colspan="4" align="center"><hr></td>
        </tr>
      <tr>
        <td colspan="4" align="center"><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="33%">&nbsp;</td>
            <td width="34%">&nbsp;</td>
            <td width="33%" align="center"><input name="button" type="submit" class="botao_cancelar" id="button" value="Imprimir"></td>
          </tr>
        </table></td>
      </tr>  
    </table></td>
  </tr>
</table>
</body>
</html>
