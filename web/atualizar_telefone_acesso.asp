
<%
Option Explicit
%>
<!--#include file="dsn.asp"-->

<% response.buffer=True%>
<%
Dim Conexao,strSQL,rs,varCodimovel,rsFoto,vFoto,strSQL2,vProposta,vimagem,vdata,vHorario,vInteresse
 
 
 dim vCidade,vBairro,vTipo,vNegociacao,vValor
 Dim vProprietario,vTelefone,vEmail,vEndereco
 dim vQuartos,vDescricao
 dim vVila,vVila2

 dim varCodCompradores
 
 dim varSucessoSenha
 
 dim varCod_telefone_acesso
 varCod_telefone_acesso = request.Querystring("varCod_telefone_acesso")
 
dim vNome,vIP,vSenha,vPermissao

dim var_telefone_acesso

var_telefone_acesso = request.form("txt_telefone_acesso")

	  													  
	
    Set Conexao = Server.CreateObject("ADODB.Connection")
	
	Conexao.Open dsn
	
	

	 
	 
	 
	 Conexao.execute"update telefone_acesso set telefone_acesso='"&var_telefone_acesso&"' where cod_telefone_acesso="&varCod_telefone_acesso
	 
	
	  response.Redirect "visualizar_telefone_acesso.asp?varSucessoSenha="&var_telefone_acesso&"&varCod_telefone_acesso="&varCod_telefone_acesso&""
     
	  
	  
   
   
   
   
   
  
   
   
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

