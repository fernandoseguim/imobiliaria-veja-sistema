<%
Option Explicit
%>
<!--#include file="dsn.asp"-->


<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,vdata,vProprietario,vEndereco,vLink_Foto,vCidade,vBairro
Dim vTipo,vAreaTotal,vAreaConstruida,vQuartos,vBanheiros,vVagas,vValor,vNegociacao,vFoto
Dim vdata2
dim vVila

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
 
 
 
  
	
	  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	
 
 
    
	
   
	dim vNome,vID,vSenha,vPermissao,vTelefone,vEmail
	
	vNome = request.form("txt_nome")
	vID = request.form("txt_id")
	 vSenha = request.form("txt_senha")
	 vPermissao = request.form("txt_permissao")
	 vTelefone = request.form("txt_telefone")
	 vEmail = request.form("txt_email")
	
	
 dim rs444Senha22,strSQL444Senha22
   
    Set rs444Senha22 = Server.CreateObject("ADODB.RecordSet")
	strSQL444Senha22 = "SELECT * FROM senha where List_Name = '"&vNome&"' or Admin_id = '"&vID&"' or Admin_pass = '"&vSenha&"'" 
	 rs444Senha22.Open strSQL444Senha22, Conexao


if not rs444Senha22.eof then

	 response.redirect "form_incluir_senha.asp?varSucesso_bairro="&"Escolha outro nome, senha ou id"&"" 
	
	
	
	end if
	Conexao.execute"Insert into senha(List_Name,Admin_ID,Admin_Pass,permissao,origem_franquia,telefone,email) values('"& vNome &"','"& vID &"','"& vSenha &"','"& vPermissao &"','"& session("vOrigem_Franquia") &"','"& vTelefone &"','"& vEmail &"')"
	 
	 dim varCidade
	 response.Redirect "form_incluir_senha.asp?varSucesso_bairro="&vNome&""
	
	  
     
	  
	  
   
   
   
   
   
  
   
   
   %>







<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>Sugestão incluída</title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
</head>

<body bgcolor="#406496" marginheight="0" marginwidth="0" topmargin="0" leftmargin="0" rightmargin="0">
<%
 
     
 
 
           
           Conexao.Close
           
           %>
<% response.flush%>
<%response.clear%>
</body>
</html>
