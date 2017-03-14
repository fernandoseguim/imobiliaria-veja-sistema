
<%
Option Explicit
%>
<!--#include file="dsn.asp"-->
<!--#include file="loggedin02.asp"-->
<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCodimovel,rsFoto,vFoto,strSQL2,vProposta,vimagem,vdata,vHorario,vInteresse
 
 
 dim vCidade,vBairro,vTipo,vNegociacao,vValor
 Dim vProprietario,vEndereco
 dim vQuartos,vDescricao
 dim vVila,vVila2

 dim varCodCompradores
 
 dim varSucessoSenha
 
 dim varCodSenha
 varCodSenha = request.Querystring("varCodSenha")
 
dim vNome,vID,vSenha,vPermissao,vTelefone,vEmail

vNome = request.form("txt_nome")

vID = request.form("txt_id")

vSenha = request.form("txt_senha")	  

vPermissao = request.form("txt_permissao")

vTelefone = request.form("txt_telefone")

vEmail = request.form("txt_email")	  													  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	

	 
	 
	 
	 Conexao.execute"update senha set List_Name='"&vNome&"',Admin_ID='"&vID&"',Admin_Pass='"&vSenha&"',permissao='"&vPermissao&"',telefone='"&vTelefone&"',email='"&vEmail&"' where ID="&varCodSenha
	 
	
	  response.Redirect "visualizar_senha.asp?varSucessoSenha="&vNome&"&varCodSenha="&varCodSenha&""
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Imóvel incluído</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#406496" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">
<%
 
     
 rs7.Close
           
           Set rs7 = Nothing
 
           
           Conexao.Close
           
           %>
<% response.flush%>
<%response.clear%>
</body>
</html>

