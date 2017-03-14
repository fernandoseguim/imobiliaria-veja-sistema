<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body>

<%

'Criando conexão com o banco de dados! 
Set Conexao = Server.CreateObject("ADODB.Connection")
Conexao.Open dsn

'Abrindo a tabela MARCAS!
Sql = "SELECT * FROM combo1 ORDER BY nome_combo1 ASC" 

Set rs = Server.CreateObject("ADODB.RecordSet")

	rs.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs.ActiveConnection = Conexao
	
	
	rs.Open sql, Conexao



'listar recordset

 if not rs303.eof then 
                  While NOT Rs303.EoF
                  
				  rs("teste01")
                  
                   Rs303.MoveNext 
                   Wend 
				   
				   else
	end if
	


'fazer uma inclusão


 Conexao.execute"Insert into permuta(Foto_imovel, Nome, Email, Telefone,endereco_vend, cidade_vend, bairro_vend,tipo_vend,descricao_vend,cidade_comp,bairro_comp,tipo_comp,descricao_comp,cod_imovel,link_imovel,data,quartos_vend,quartos_comp,valor_vend,valor_comp,atendimento,data_atualizacao,vila_vend,vila_comp,vagas_vend,vagas_comp) values( '"& vimagem &"','"& vProprietario_vend &"','"& vEmail_vend &"','"& vTelefone_vend &"','"& vEndereco_vend &"','"& vCidade2_vend &"','"& vBairro2_vend &"','"& vTipo_vend &"','"& vOBS_imovel_vend &"','"& vCidade2 &"','"& vBairro2 &"','"& vTipo &"','"& vDescricao &"','"& varCodImovel444 &"','"& "00" &"','"& now() &"','"& vQuartos_vend &"','"& vQuartos &"','"& int(vValor_vend) &"','"& int(vStage22) &"','"& vAtendimento &"','"& now() &"','"& vVila2_vend &"','"& vVila2 &"','"& vVagas_vend &"','"& vVagas &"')"
 
 
 'fazer uma inclusão

Conexao.execute"Delete from imoveis where cod_imovel="&VarCod_imovel
	
	
	'Fazer uma atualização
	
	 Conexao.execute"update imoveis set proprietario='"&vProprietario_vend&"',telefone='"&vTelefone_vend&"',email='"&vEmail_vend&"' where cod_imovel="&varCod_imovel006

' Envio de Email


Set Mailer = Server.CreateObject("SMTPsvg.Mailer")
Mailer.ContentType = "text/html"
Mailer.RemoteHost = "smtp.testeaspl.com.br" 
Mailer.FromName = "TESTE - ASPL"
Mailer.FromAddress = "contato@asp.com.br"
Mailer.AddRecipient rsquery("nome"),rsquery("email") 
Mailer.Subject=request.form("assunto")
Mailer.Bodytext = "Caro <b>" & rsquery("nome") & ",</b>" & request.form("texto")
x = Mailer.SendMail 







%>

<%
'--------------------Listar corretores-----------------------------

 dim rs444Atendimento,strSQL444Atendimento
   
    Set rs444Atendimento = Server.CreateObject("ADODB.RecordSet")
	strSQL444Atendimento = "SELECT senha.list_name,senha.admin_id,senha.admin_pass,senha.admin_email,senha.from_email,senha.url_dir,senha.url_home,senha.component,senha.smtp,senha.permissao,senha.id  FROM senha  ORDER BY id" 
	
	
	rs444Atendimento.CursorLocation = 3
    rs444Atendimento.CursorType = 3

    rs444Atendimento.ActiveConnection = Conexao
	
	
	
	
	
	 rs444Atendimento.Open strSQL444Atendimento, Conexao



dim varAtendimento
varAtendimento = request.querystring("txt_atendimento")
if varAtendimento = "" then
varAtendimento = request.querystring("varAtendimento")
end if

session("varAtendimento") = varAtendimento



'-------------------------Conversão de datas---------------------

"select data, nom_crianca, pai, mae, codigo from criancas where data between convert(smalldatetime,'" + inicial + "', 103) and convert(smalldatetime,'" + final + "', 103)






'-------------------------------------------------------------------

%>

</body>
</html>
