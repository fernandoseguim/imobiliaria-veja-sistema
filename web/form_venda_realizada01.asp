<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<!--#include file="style_imoveis.asp"-->



<%
dim Conexao
dim strSQL
dim rs
dim varCod_imovel

letra = "#FFFFFF"

 Set Conexao = Server.CreateObject("ADODB.Connection")
 Set rs = Server.CreateObject("ADODB.Recordset")
 
 varCod_imovel = request.QueryString("varCod_imovel")
 	
	
	strSQL = "SELECT *  FROM imoveis where cod_imovel="&varCod_imovel

   Conexao.Open dsn

rs.CursorLocation = 3
rs.CursorType = 3

        rs.Open strSQL, Conexao 
	

%>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="<%=escuro%>">
<div align="center"><font color="#FFFFFF" size="4" face="Verdana, Arial, Helvetica, sans-serif">Venda realizada</font></div>
<%
dim varSucessoVenda

varSucessoVenda = request.QueryString("varSucessoVenda")



%>


 <div align="center"><%
	if varSucessoVenda = "" then
	response.Write varSucessoVenda
	%>
        <%else%>
        <font color="#FFFFFF" size="1" face="Verdana, Arial, Helvetica, sans-serif"><%=varSucessoVenda%> 
        foi atualizado com sucesso.</font> 
        <% end if %>
		</div>




<form name="doublecombo"   method="post" action="atualizar_venda_realizada01.asp?varCod_imovel=<%=varCod_imovel%>">

<table width="1000" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td><table width="1000">
  <tr>
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Propriet&aacute;rio 
      do im&oacute;vel</font></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">data 
      de nascimento do propriet&aacute;rio</font></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Comprador 
      do im&oacute;vel</font></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
      de nascimento do comprador</font></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">N&uacute;mero 
      da venda</font></td>
  </tr>
  <tr>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_proprietario_venda" type="text" class="inputBox" id="txt_proprietario_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("proprietario_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_data_nasc_proprietario_venda" type="text" class="inputBox" id="txt_data_nasc_proprietario_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("data_nasc_proprietario_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_comprador_venda" type="text" class="inputBox" id="txt_comprador_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("comprador_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_data_nasc_comprador_venda" type="text" class="inputBox" id="txt_data_nasc_comprador_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("data_nasc_comprador_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_numero_venda" type="text" class="inputBox" id="txt_numero_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("numero_venda")%>" size="38" maxlength="50" align="left"></td>
  </tr>
  <tr>
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Data 
      da venda</font></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Corretor 
      da venda</font></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Captador 
      da venda</font></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Forma 
      de pagamento utilizada</font></td>
    <td width="10">&nbsp;</td>
          
    <td width="192" bgcolor="<%=escuro%>">&nbsp;</td>
  </tr>
  <tr>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_data_venda" type="text" class="inputBox" id="txt_data_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("data_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_corretor_venda" type="text" class="inputBox" id="txt_corretor_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("corretor_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_captador_venda" type="text" class="inputBox" id="txt_captador_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("captador_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_forma_pagamento_venda" type="text" class="inputBox" id="txt_forma_pagamento_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("forma_pagamento_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">&nbsp;</td>
  </tr>
  <tr>
    <td width="192" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
      da venda</font></td>
    <td width="10">&nbsp;</td>
          
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Valor 
      da comiss&atilde;o de venda</font></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Custo 
      da venda</font></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Custo 
      da comiss&atilde;o do corretor</font></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Custo 
      da comiss&atilde;o do captador</font></td>
  </tr>
  <tr>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_valor_venda" type="text" class="inputBox" id="txt_valor_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("valor_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_valor_comissao_venda" type="text" class="inputBox" id="txt_valor_comissao_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("valor_comissao_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_custo_venda" type="text" class="inputBox" id="txt_custo_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("custo_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_custo_corretor_venda" type="text" class="inputBox" id="txt_custo_corretor_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("custo_corretor_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_custo_captador_venda" type="text" class="inputBox" id="txt_custo_captador_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("custo_captador_venda")%>" size="38" maxlength="50" align="left"></td>
  </tr>
  <tr>
    <td width="192" bgcolor="<%=escuro%>" ><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Custo 
      da comiss&atilde;o do gerente de vendas</font></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Custo 
      com documenta&ccedil;&atilde;o</font></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Custos 
      extras</font></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=escuro%>"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Lucro 
      l&iacute;quido</font></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=escuro%>">&nbsp;</td>
  </tr>
  <tr>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_custo_gerente_venda" type="text" class="inputBox" id="txt_custo_gerente_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("custo_gerente_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_custo_documentacao_venda" type="text" class="inputBox" id="txt_custo_documentacao_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("custo_documentacao_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_custo_extra_venda" type="text" class="inputBox" id="txt_custo_extra_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("custo_extra_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><input name="txt_lucro_liquido_venda" type="text" class="inputBox" id="txt_lucro_liquido_venda" style="HEIGHT: 18px; WIDTH: 190px; background:<%=medio%>;" value="<%=rs("lucro_liquido_venda")%>" size="38" maxlength="50" align="left"></td>
    <td width="10">&nbsp;</td>
    <td width="192" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;">&nbsp;</td>
  </tr>
  
  
 
  
   </table></td>
  </tr>
  <tr>
    <td><table width="1000" border="0" cellspacing="0" cellpadding="0">
  <tr>
          <td width="1000" height="40"> 
            <div align="center"><font color="<%=letra%>" size="1" face="Verdana, Arial, Helvetica, sans-serif">Abaixo 
              escreva a origem e o hist&oacute;rico da venda</font></div></td>
  </tr>
  <tr>
          <td height="120" bgcolor="<%=medio%>" style="border:1px solid #FFFFFF;"><textarea name="txt_origem_venda" class="inputBox" id="txt_origem_venda" style="HEIGHT: 100px; WIDTH: 1000px; background:<%=medio%>" onKeyPress="return limitfield(this, 800)"><%=rs("origem_venda")%></textarea></td>
  </tr>
  <tr>
          <td height="20"><div align="right">
                <input name="image" type="image"  src="bt_enviar0011.jpg" width="145" height="18" border="0">
                <a href="javascript:document.forms.doublecombo.reset()"><img src="bt_apagar0011.jpg" width="145" height="18" border="0"></a></div></td>
  </tr>
</table></td>
  </tr>
</table>

</form>
<%

            rs.Close           
		   
           Set rs = Nothing
		   
		   Conexao.close
		   
		  Set Conexao = Nothing
		   


%>


</body>
</html>
