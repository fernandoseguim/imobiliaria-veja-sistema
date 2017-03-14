
<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>
<%


'----------------------Declarar variáveis-------------------------------------

dim vProprietario_venda
dim vComprador_venda
dim vData_nasc_proprietario_venda
dim vData_nasc_comprador_venda
dim vData_venda
dim vValor_venda
dim vCorretor_venda
dim vCaptador_venda
dim vValor_comissao_venda
dim vForma_pagamento_venda
dim vOrigem_venda
dim vNumero_venda
dim vCusto_venda
dim vCusto_corretor_venda
dim vCusto_captador_venda
dim vCusto_gerente_venda
dim vCusto_documentacao_venda
dim vCusto_extra_venda
dim vLucro_liquido_venda

 
 
 
 '---------------vamos pegar as informações do formulário-----------------------
 
 vProprietario_venda = request.form("txt_proprietario_venda")
 vComprador_venda = request.form("txt_comprador_venda")
 vData_nasc_proprietario_venda = request.form("txt_data_nasc_proprietario_venda")
 vData_nasc_comprador_venda = request.form("txt_data_nasc_comprador_venda")
 vData_venda = request.form("txt_data_venda")
 vValor_venda = request.form("txt_valor_venda")
 vCorretor_venda = request.form("txt_corretor_venda")
 vCaptador_venda = request.form("txt_captador_venda")
 vValor_comissao_venda = request.form("txt_valor_comissao_venda")
 vForma_pagamento_venda = request.form("txt_forma_pagamento_venda")
 vOrigem_venda = request.form("txt_origem_venda")
 vNumero_venda = request.form("txt_numero_venda")
 vCusto_venda = request.form("txt_custo_venda")
 vCusto_corretor_venda = request.form("txt_custo_corretor_venda")
 vCusto_captador_venda = request.form("txt_custo_captador_venda")
 vCusto_gerente_venda = request.form("txt_custo_gerente_venda")
 vCusto_documentacao_venda = request.form("txt_custo_documentacao_venda")
 vCusto_extra_venda = request.form("txt_custo_extra_venda")
 vLucro_liquido_venda = request.form("txt_lucro_liquido_venda")
 
 
 
 
 
 
 '--------------------------------------------------------------------------------
 
 
 
 
 
 dim varCod_imovel
 varCod_imovel = request.Querystring("varCod_imovel")
 





	  dim Conexao													  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	

	 
	 
	 
	 Conexao.execute"update imoveis set proprietario_venda='"&vProprietario_venda&"',comprador_venda='"&vComprador_venda&"',data_nasc_proprietario_venda='"&vData_nasc_proprietario_venda&"',data_nasc_comprador_venda='"&vData_nasc_comprador_venda&"',data_venda='"&vData_venda&"',valor_venda='"&vValor_venda&"',corretor_venda='"&vCorretor_venda&"',Captador_venda='"&vCaptador_venda&"',valor_comissao_venda='"&vValor_comissao_venda&"',forma_pagamento_venda='"&vForma_pagamento_venda&"',origem_venda='"&vOrigem_venda&"',numero_venda='"&vNumero_venda&"',custo_venda='"&vCusto_venda&"',custo_corretor_venda='"&vCusto_corretor_venda&"',custo_captador_venda='"&vCusto_captador_venda&"',custo_gerente_venda='"&vCusto_gerente_venda&"',custo_documentacao_venda='"&vCusto_documentacao_venda&"',custo_extra_venda='"&vCusto_extra_venda&"',lucro_liquido_venda='"&vLucro_liquido_venda&"' where cod_imovel="&varCod_imovel	 
	
	  response.Redirect "form_venda_realizada01.asp?varSucessoVenda="&vProprietario_venda&"&varCod_imovel="&varCod_imovel&""
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Imóvel incluído</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#406496" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">





<table width="590" height="462" cellpadding="0" cellspacing="0" bgcolor="#406496">

<tr>
<td width="590" height="48"><img src="top_resultado.jpg"></img></td>
</tr>
<tr>
<td width="590" height="105" ></td>
</tr>
<tr>
<td width="590" height="156" >

<table cellspacing="0" cellpadding="0">
<tr>
<td width="217" height="156" ></td>    <td width="202" height="156" ><img src="sorriso_proposta.jpg" border="0"></img></td>   <td width="217" height="156" ></td>
</tr>

</table>



</td>
</tr>
<tr>
<td width="590" height="117" ></td>
</tr>


<tr>
    <td width="590" height="36" ></img></td>

</tr>


</table>







 
 <%
 
     
 rs7.Close
           
           Set rs7 = Nothing
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>

