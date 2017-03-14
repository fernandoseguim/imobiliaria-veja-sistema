

<%
Option Explicit
%>
<!--#include file="cores02.asp"-->
<!--#include file="dsn.asp"-->
<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCodImovel,rsFoto,vFoto,vNome,vTelefone,vEmail,strSQL2,vProposta,vimagem,vdata,vHorario,vNegociacao
 
 varCodImovel = request.QueryString("varCodImovel")
   
  
   dim varCod_comprador02
   
  varCod_comprador02 = request.QueryString("varCod_comprador02")
   
   Dim vdata2

vdata2 = now()

if len(vdata2) = 17 then
 vdata = left(now(),9)
 end if
 
 if len(vdata2) = 18 then
 vdata = left(now(),10)
 end if
 
 if len(vdata2) = 19 then
 vdata = left(now(),11)
 end if
  vNome=request.form("txtNome")
      vTelefone=request.form("txtTelefone")
      vEmail=request.form("txtEmail")
	  
	  if vEmail = "" then
	  vEmail = "não informado"
	  end if
	  
	  vProposta=request.form("txtProposta")                                                           
		vHorario=request.form("txtHorario")
		vNegociacao=request.form("txtNegociacao")													  
	Set rs = Server.CreateObject("ADODB.RecordSet")
    Set Conexao = Server.CreateObject("ADODB.Connection")
	'strSQL = "SELECT * FROM imoveis  Where cod_imovel = "&varCodImovel
	dim strSQLComprador
	
	strSQLComprador = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento,compradores.historico_atual01,compradores.historico_atual02,compradores.historico_atual03,compradores.historico_atual04,compradores.historico_atual05,compradores.historico_atual06,compradores.historico_quem01,compradores.historico_quem02,compradores.historico_quem03,compradores.historico_quem04,compradores.historico_quem05,compradores.historico_quem06,compradores.ocupacao_hist,compradores.endereco_hist,compradores.valor_hist,compradores.quartos_hist,compradores.vagas_hist,compradores.suites_hist,compradores.piscina_hist,compradores.area_total_hist,compradores.area_construida_hist,compradores.edicula_hist,compradores.condominio_hist  FROM compradores where cod_compradores="&varCod_comprador02
  
  
	
	 Conexao.Open dsn
	 
	 
	 
   
RS.CursorLocation = 3
RS.CursorType = 3

        rs.Open strSQLComprador, Conexao 
	dim AtendimentoComprador
	if not rs.eof then
	AtendimentoComprador=rs("atendimento")
	else
	AtendimentoComprador = "internet"
	end if
	
	
	'Conexao.execute"Insert into proposta(Foto_proposta, Nome_proposta, Email_proposta, Telefone_proposta, Proposta_proposta, data_proposta,horario_proposta,interesse_proposta,cod_imovel_proposta,origem_franquia) values( '"& vimagem &"','"& vNome &"','"& vEmail &"','"& vTelefone &"','"& vProposta &"','"& now() &"','"& vHorario &"','"& vNegociacao &"','"& varCodImovel &"','"&session("vOrigem_Franquia")&"')" 
	 
	  Conexao.execute"Insert into email(nome, telefone, email ,assunto,mensagem,data,cod_imovel,atendimento,origem,origem_franquia) values( '"& vNome &"','"& vTelefone &"','"& vEmail &"','"& "Quero um comprador para o meu imóvel" &"','"& vProposta &"','"& now() &"','"& "0" &"','"& AtendimentoComprador &"','"& "Busca de Compradores" &"','"& session("vOrigem_Franquia") &"')" 
	
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
   
   
  	  '--------------------------Se não tiver cadastrar-----------
 ' dim rs444VerificaConta2,strSQL444VerificaConta2
   
  '  Set rs444VerificaConta2 = Server.CreateObject("ADODB.RecordSet")
	'strSQL444VerificaConta2 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like'%"&vTelefone&"%' or telefone02 like'%"&vTelefone&"%' or telefone03 like'%"&vTelefone&"%'" 
	
	
	'rs444VerificaConta2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

'rs444VerificaConta2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

'rs444VerificaConta2.ActiveConnection = Conexao
	
	
	
	' rs444VerificaConta2.Open strSQL444VerificaConta2, Conexao
	 
   
  
   
   
	
	' dim vOrigem
	 
	' vOrigem = "Busca de comprador"
	 
	' dim vAtendimento
	 
	' if  not rs444VerificaConta2.eof then
	 
	' vAtendimento = rs444VerificaConta2("atendimento")
	' else
	' vAtendimento = "não informado"
	' end if
	 
	 
	
	  
   
   
   
  'Conexao.execute"Insert into email(nome, telefone, email ,assunto,mensagem,data,cod_imovel,atendimento,origem) values( '"& vNome &"','"& vTelefone &"','"& vEmail &"','"& vNegociacao &"','"& vProposta &"','"& now() &"','"& varCod_comprador02 &"','"& vAtendimento &"','"& vOrigem &"')" 
	 
	 
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Proposta pelo imóvel incluída</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="<%=escuro%>" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">

<table width="590" height="462" cellpadding="0" cellspacing="0">

<tr>
    <td width="590" height="48"><img src="top_resultado02.jpg" width="590" height="48"></img></td>
</tr>
<tr>
    <td width="590" height="105" bgcolor="<%=escuro%>" ></td>
</tr>
<tr>
<td width="590" height="156" >

<table cellspacing="0" cellpadding="0">
<tr>
          <td width="217" height="156" bgcolor="<%=escuro%>" ></td>    
          <td width="202" height="156" bgcolor="<%=escuro%>" ></img> <div align="center"><font color="#FFFFFF"><strong>Sua 
              mensagem foi inclu&iacute;da com sucesso!!</strong></font></div></td>   
          <td width="217" height="156" bgcolor="<%=escuro%>" ></td>
</tr>

</table>



</td>
</tr>
<tr>
    <td width="590" height="117" bgcolor="<%=escuro%>" ></td>
</tr>


<tr>
    <td width="590" height="36" bgcolor="<%=escuro%>" ></img></td>

</tr>


</table>












 
 <%
 
     
 
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>


