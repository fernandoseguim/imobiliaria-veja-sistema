<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>
<%


if session("vOrigem_Franquia") = "" then
session("vOrigem_Franquia") = "São Bernardo"
end if


Dim Conexao,strSQL,rs,vdata,vNome,vEmail,vAssunto,vMensagem,vTelefone
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
    vEmail=request.form("txtEmail") 
	vTelefone = request.form("txtTelefone")
	
	
	
	
	if vEmail = "" then
	vEmail = "não informado"
	end if 
   
    vAssunto=request.form("txtAssunto")   
     vMensagem=request.form("txtMensagem")                                                   
															  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	 
	
	
	 
	'--------------------Verificar se já tem cadastro-------------- 
	 
	  '--------------------------Se não tiver cadastrar-----------
  dim rs444VerificaConta2,strSQL444VerificaConta2
   
    Set rs444VerificaConta2 = Server.CreateObject("ADODB.RecordSet")
	strSQL444VerificaConta2 = "SELECT compradores.cod_compradores,compradores.nome,compradores.telefone,compradores.email,compradores.endereco,compradores.cidade,compradores.bairro,compradores.tipo,compradores.quartos,compradores.negociacao,compradores.valor,compradores.descricao,compradores.data,compradores.data_atualizacao,compradores.atendimento,compradores.vila,compradores.vagas,compradores.ocupacao,compradores.standby,compradores.cod_imovel,compradores.cod_permuta,compradores.acessos,compradores.descricao_confi,compradores.origem,compradores.dataLastEmail,compradores.textoLastEmail,compradores.responsavel_cadastramento,compradores.data_ultimo_acesso,compradores.data_futuro_contato,compradores.assunto_futuro_contato,compradores.melhor_horario_visita,compradores.telefone02,compradores.telefone03,compradores.data_ligar_urgente,compradores.assunto_ligar_urgente,compradores.quem_atualizou  FROM compradores where telefone like'%"&vTelefone&"%' or telefone02 like'%"&vTelefone&"%' or telefone03 like'%"&vTelefone&"%'" 
	
	
	rs444VerificaConta2.CursorLocation = 3
'a propriedade CursorLocation do objeto recordset indica onde o cursor é criado
'se no cliente ou no servidor.

rs444VerificaConta2.CursorType = 3
'indica o tipo de cursor utilizado, se somente de leitura ou se de leitura e gravação.

rs444VerificaConta2.ActiveConnection = Conexao
	
	
	
	 rs444VerificaConta2.Open strSQL444VerificaConta2, Conexao
	 
    if rs444VerificaConta2.eof then
	
	
	
	Conexao.execute"Insert into compradores(nome,telefone,email,cidade,bairro,tipo,quartos,negociacao,valor,data,descricao,atendimento,data_atualizacao,vila,vagas,ocupacao,standby,cod_imovel,cod_permuta,acessos,descricao_confi,origem,origem_franquia) values( '"& vNome &"','"& vTelefone &"','"& vEmail &"','"& "não informado" &"','"& "não informado" &"','"& "não informado" &"','"& "0" &"','"& "não informado" &"','"& "0" &"','"& now() &"','"& "não informado" &"','"& "internet" &"','"& now() &"','"& "não informado" &"','"& "0" &"','"& "não informado" &"','"& "excluido" &"','"& "0" &"','"& "0" &"','"& "0" &"','"& "Não informado" &"','"& "internet" &"','"& session("vOrigem_Franquia") &"')"
	
	
	end if 
	
	
	 dim vOrigem
	 
	 vOrigem = "Email"
	 
	 dim vAtendimento
	 
	 if  not rs444VerificaConta2.eof then
	 
	 vAtendimento = rs444VerificaConta2("atendimento")
	 else
	 vAtendimento = "não informado"
	 end if
	 
	 
	
	 
	 
	 
	 
	 '--------------------------------------------------------------
	 
	  Conexao.execute"Insert into email(nome, telefone, email ,assunto,mensagem,data,cod_imovel,atendimento,origem,origem_franquia) values( '"& vNome &"','"& vTelefone &"','"& vEmail &"','"& vAssunto &"','"& vMensagem &"','"& now() &"','"& "0" &"','"& vAtendimento &"','"& vOrigem &"','"& session("vOrigem_Franquia") &"')" 
	dim varSucesso
	
	response.Redirect "form_enviar_email.asp?varSucesso="&"Mensagem enviada com sucesso"&""
	  
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Sugestão incluída</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#406496" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">





<table width="590" height="462" cellpadding="0" cellspacing="0" bgcolor="#406496">

<tr>
<td width="590" height="48"  ><img src="top_resultado.jpg"></img></td>
</tr>
<tr>
<td width="590" height="105" ></td>
</tr>
<tr>
<td width="590" height="156" >

<table cellspacing="0" cellpadding="0">
<tr>
<td width="217" height="156" ></td>    
          <td width="202" height="156"><img src="sorriso_email.jpg" width="202" height="156" border="0"></img></td>   
          <td width="217" height="156" ></td>
</tr>

</table>



</td>
</tr>
<tr>
<td width="590" height="117" ></td>
</tr>


<tr>
    <td width="590" height="36"></img></td>

</tr>


</table>







 
 <%
 
     
 
 
           
           Conexao.Close
           
           %>
		   
		   
<% response.flush%>
	   <%response.clear%>
</body>
</html>
