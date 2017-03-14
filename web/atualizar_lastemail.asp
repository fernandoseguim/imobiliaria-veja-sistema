<!--#include file="dsn.asp"-->
<!--#include file="cores.asp"-->
<%
dim varDe,varPara,varAssunto,varMensagem

dim vNome
dim vTelefone
dim vEmail

dim Conexao

Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn



dim varCodCompradores

varCodCompradores = request.querystring("varCodCompradores")

dim strSQL001

	strSQL001 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou,compradores.obs_quartos,compradores.obs_vagas,compradores.suites,compradores.obs_suites,compradores.salao_de_festas,compradores.obs_salao_de_festas,compradores.salao_de_jogos,compradores.obs_salao_de_jogos,compradores.piscina,compradores.obs_piscina,compradores.andares_edificio,compradores.obs_andares_edificio,compradores.edicula,compradores.obs_edicula,compradores.quintal,compradores.obs_quintal,compradores.banheiros,compradores.obs_banheiros,compradores.entrada_lateral,compradores.obs_entrada_lateral,compradores.churrasqueira,compradores.obs_churrasqueira,compradores.quadras,compradores.obs_quadras,compradores.portaria,compradores.obs_portaria,compradores.quantidade_elevadores,compradores.quantidade_elevadores,compradores.obs_quantidade_elevadores,compradores.area_total,compradores.area_construida,compradores.condominio,compradores.condicoes_pagamento,pergunta,tarja02,data01_tarja02,data02_tarja02,compradores.obs_forma_pagamento  FROM compradores where cod_compradores="&varCodCompradores
	
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
	vNome = rs001("nome")
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
	 
	 
	 
	 Conexao.execute"update compradores set dataLastEmail='"&now()&"',textoLastEmail='"&varMensagem&"' where cod_compradores="&varCodCompradores
	 
	

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

<%
conexao.close
set conexao = nothing
%>
<!--#include file="dsn2.asp"-->