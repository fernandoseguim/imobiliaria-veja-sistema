<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<%
dim varDe,varPara,varAssunto,varMensagem

dim varCod_imovel
 
dim  strSQL001

dim vNome
dim vTelefone
dim vEmail
dim Conexao

 Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn

varCod_imovel = request.querystring("varCod_imovel")

strSQL001 = "SELECT imoveis.cod_imovel,imoveis.foto_grande,imoveis.cidade,imoveis.bairro,imoveis.tipo,imoveis.area_total,imoveis.area_construida,imoveis.quartos,imoveis.banheiros,imoveis.vagas,imoveis.negociacao,imoveis.valor,imoveis.link_foto,imoveis.proprietario,imoveis.endereco,imoveis.data,imoveis.telefone,imoveis.email,imoveis.obs_imovel,imoveis.obs_proprietario,imoveis.foto_pequena,imoveis.presenca_primeira,imoveis.titulo_anuncio,imoveis.texto_anuncio,imoveis.foto_grande1,imoveis.foto_grande2,imoveis.foto_grande3,imoveis.foto_grande4,imoveis.foto_grande5,imoveis.StandBy,imoveis.foto_gigante,imoveis.ocupacao,imoveis.data_atualizacao,imoveis.captacao,imoveis.vila,imoveis.qualidade,imoveis.foto_grande6,imoveis.foto_grande7,imoveis.foto_grande8,imoveis.foto_grande9,imoveis.foto_grande10,imoveis.cod_permuta,imoveis.cod_comprador,imoveis.condominio,imoveis.placa,imoveis.dataLastEmail,imoveis.textoLastEmail,imoveis.data_futuro_contato,imoveis.assunto_futuro_contato,imoveis.telefone02,imoveis.telefone03,imoveis.suites,imoveis.chaves_do_imovel,imoveis.melhor_horario_visita,imoveis.imovel_em_negociacao,imoveis.metros_de_frente,imoveis.metros_de_fundo,imoveis.metros_lateral_esquerda,imoveis.metros_lateral_direita,imoveis.data_captacao,imoveis.origem_captacao,imoveis.responsavel_cadastramento,imoveis.data_ultimo_acesso,imoveis.saldo_devedor,imoveis.ja_pago_devedor,imoveis.devendo_devedor,imoveis.quem_atualizou,imoveis.obs_quartos,imoveis.obs_vagas,imoveis.obs_banheiros,imoveis.obs_edicula,imoveis.obs_entrada_lateral,imoveis.obs_salao_de_festas,imoveis.obs_salao_de_jogos,imoveis.obs_churrasqueira,imoveis.obs_piscina,imoveis.obs_quintal,imoveis.obs_quadras,imoveis.obs_andares_edificio,imoveis.obs_quantidade_elevadores,imoveis.obs_portaria,imoveis.obs_suites,imoveis.salao_de_festas,imoveis.piscina,imoveis.andares_edificio,imoveis.edicula,imoveis.salao_de_jogos,imoveis.quintal,imoveis.quantidade_elevadores,imoveis.entrada_lateral,imoveis.churrasqueira,imoveis.quadras,imoveis.portaria,imoveis.valor_iptu,imoveis.valor_outros,imoveis.nome_edificio,imoveis.conseguiu_proposta,imoveis.quem_tirou_foto,imoveis.rateio,imoveis.pergunta,imoveis.tarja02,imoveis.data01_tarja02,imoveis.data02_tarja02,imoveis.cliques_no_imovel,imoveis.obs_forma_pagamento  FROM imoveis where cod_imovel="&varCod_imovel


dim rs001
	Set rs001 = Server.CreateObject("ADODB.RecordSet")

	rs001.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs001.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs001.ActiveConnection = Conexao
	
	
	rs001.Open strSQL001, Conexao
	
	if not rs001.eof then
	vNome = rs001("proprietario")
	vTelefone = rs001("telefone")
	vEmail = rs001("email")
	else
	vNome = "não informado"
	vTelefone = "não informado"
	vEmail = "não informado"
	end if



varDe = request.form("txtDe")
varPara = request.form("txtPara")
varAssunto = request.form("txtAssunto")
varMensagem = request.form("txtMensagem")


	
	

	 
	 
	 
	  Conexao.execute"Insert into email_enviado(nome, telefone, email ,assunto,mensagem,data,atendimento,de,para,origem_franquia) values( '"& vNome &"','"& vTelefone &"','"& vEmail &"','"& varAssunto &"','"& varMensagem &"','"& now() &"','"& session("nome_id") &"','"& varDe &"','"& varPara &"','"& session("vOrigem_Franquia") &"')" 
	 
	 
	 
	 Conexao.execute"update imoveis set dataLastEmail='"&now()&"',textoLastEmail='"&varMensagem&"' where cod_imovel="&varCod_imovel
	 
	
 
%>

<html>
<body bgcolor="<%=escuro%>">
<%
Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
Mailer.FromName = "Imobiliária Veja"
Mailer.FromAddress= varDe
Mailer.RemoteHost = "smtp2.locaweb.com.br"
Mailer.AddRecipient "Imobiliária Veja", varPara
Mailer.Subject = varAssunto
Mailer.BodyText = varMensagem
If Mailer.SendMail Then
Response.Write "<br><br><center><strong><font color='#FFFFFF' size='2' face='Verdana, Arial, Helvetica, sans-serif'>Mensagem enviada!!</font> </strong></center>"
Else
Response.Write "Erro " & Mailer.Response
End If

Set Mailer = Nothing
%>
</body>
</html>


<!--#include file="dsn2.asp"-->
<%

conexao.close
set conexao = nothing

%>