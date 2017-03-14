<% response.Buffer = true %>
<!--#include file="dsn.asp"-->
<!--#include file="cores02.asp"-->


<%

if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "São Bernardo"
end if


dim Conexao

 Set Conexao = Server.CreateObject("ADODB.Connection")
	
	
	dim varValor
	dim varCodImovel
	
	varValor = request.QueryString("varValor")
	varCodImovel = request.QueryString("varCodImovel")
	 
   Conexao.Open dsn



dim strSQLFinancia
 dim rsFinancia

strSQLFinancia = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta FROM compradores where (telefone like'%"&  session("telefone")&"%' or telefone02 like'%"&  session("telefone")&"%' or telefone03 like'%"&  session("telefone")&"%') order by cod_compradores DESC"
		
	
Set rsFinancia = Server.CreateObject("ADODB.RecordSet")

	rsFinancia.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rsFinancia.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rsFinancia.ActiveConnection = Conexao
	
	
	rsFinancia.Open strSQLFinancia, Conexao



 if not rsFinancia.eof  then 
				    
                   
                    'Conexao.execute"update imoveis set imovel_em_negociacao='"&"não informado"&"' where imovel_em_negociacao IS NULL" 
                   
				    Conexao.execute"update  compradores set condicoes_pagamento='"&"Financiamento"&"',Descricao='"&rsFinancia("descricao")&" Este comprador quer fazer financiamento."&"' where  cod_compradores="&rsFinancia("cod_compradores")
                   
				 
end if


	'-------------------------------
   rsFinancia.close
  
  
  set rsFinancia = nothing
 '-----------------------------

Conexao.execute"Insert into financiamentos(nome,telefone,email,valor,cod_imovel,data,origem_franquia) values( '"& session("nome") &"','"& session("telefone") &"','"& session("email") &"','"& varValor &"','"&  varCodImovel &"','"& now() &"','"& session("vOrigem_Franquia") &"')"
	

%>



<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Simulador de financiamento</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">

<SCRIPT LANGUAGE=JAVASCRIT TYPE="text/javascript">

function newWindow99(abrejanela99) {
   openWindow99 = window.open(abrejanela99,'openWin99','width=462,height=473,resizable=yes,left=100,scrollbars=yes')
   openWindow99.focus( )
   }

</SCRIPT>



</head>

<body bgcolor="#f7ecbf">
<table width="300" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td height="40" bgcolor="#988e47"> 
      <div align="center"><font color="#FFFFFF" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong>Clique 
        abaixo no banco de sua prefer&ecirc;ncia:</strong></font></div></td>
  </tr>
  
  
  
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="40" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
      <div align="center"><font color="<%=letra%>" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow99('http://www8.caixa.gov.br/siopiinternet/simulaOperacaoInternet.do?method=inicializarCasoUso')" style="color:#9d9249;text-decoration:none;">Caixa 
        carta FGTS</a></strong></font></div></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="40" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
      <div align="center"><font color="<%=letra%>" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow99('https://ww3.itau.com.br/imobline/pre/simuladores/index.aspx?imob_tipoBkl=&ident_bkl=pre')" style="color:#9d9249;text-decoration:none;">Ita&uacute;</a></strong></font></div></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="40" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
      <div align="center"><font color="<%=letra%>" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow99('http://www.shopcredit.com.br/shopcredit/default.asp?pag=br/stsm/stsmtipocredimob.asp?layout=F&strPerfil=8&simpj=N')" style="color:#9d9249;text-decoration:none;">Bradesco</a></strong></font></div></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="40" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
      <div align="center"><font color="<%=letra%>" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow99('http://www.hsbc.com.br/para-voce/emprestimos-financiamentos/credito-imobiliario.shtml')" style="color:#9d9249;text-decoration:none;">HSBC</a></strong></font></div></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="40" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
      <div align="center"><font color="<%=letra%>" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow99('http://www.unibanco.com.br/simuladoresubb/asp/pfisica/simuladores/imobiliarionovo/asp/index.asp')" style="color:#9d9249;text-decoration:none;">Unibanco</a></strong></font></div></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="40" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
      <div align="center"><font color="<%=letra%>" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow99('http://www.bancoreal.com.br/index_internas.htm?sUrl=http://www.bancoreal.com.br/creditoimobiliario/quero_contratar/tpl_quero_simulador.shtm')" style="color:#9d9249;text-decoration:none;">Real</a></strong></font></div></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
  <tr> 
    <td height="40" bgcolor="<%=claro%>" style="border:1px solid #FFFFFF;"> 
      <div align="center"><font color="<%=letra%>" size="2" face="Verdana, Arial, Helvetica, sans-serif"><strong><a href="javascript:newWindow99('http://www.santander.com.br/portal/gsb/script/templates/GCMRequest.do?page=496&entryID=2883')" style="color:#9d9249;text-decoration:none;">Santander</a></strong></font></div></td>
  </tr>
  <tr> 
    <td>&nbsp;</td>
  </tr>
</table>

</body>
</html>
